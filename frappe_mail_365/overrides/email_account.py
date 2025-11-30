import email.utils
import requests
import frappe
from frappe import _
from frappe.email.doctype.email_account.email_account import EmailAccount


GRAPH_API_BASE = "https://graph.microsoft.com/v1.0"
GRAPH_INBOX_URL = f"{GRAPH_API_BASE}/me/mailFolders/inbox/messages"
GRAPH_SEND_URL = f"{GRAPH_API_BASE}/me/sendMail"
REQUEST_TIMEOUT = 30


class Mail365EmailAccount(EmailAccount):
    """Email Account with Microsoft 365 Graph API support."""

    def validate(self):
        if self._is_graph_api_enabled():
            self._validate_graph_settings()
        super().validate()

    def validate_smtp_conn(self):
        if self._is_graph_api_enabled():
            return None
        return super().validate_smtp_conn()

    def validate_imap(self):
        if self._is_graph_api_enabled():
            return None
        return super().validate_imap()

    def get_incoming_server(self, *args, **kwargs):
        if self._is_graph_api_enabled():
            return None
        return super().get_incoming_server(*args, **kwargs)

    def _is_graph_api_enabled(self):
        return bool(getattr(self, "use_graph_api", 0))

    def _validate_graph_settings(self):
        if self.auth_method != "OAuth":
            frappe.throw(_("Microsoft 365 Graph API requires OAuth authentication."))

        if not self.connected_app:
            frappe.throw(_("Please select a Connected App for Microsoft 365."))

    def _get_oauth_token(self):
        try:
            if self.auth_method != "OAuth" or not self.connected_app:
                return None

            connected_app = frappe.get_doc("Connected App", self.connected_app)
            return connected_app.get_active_token(self.connected_user)

        except Exception:
            frappe.log_error(
                title="Mail 365: Token Error",
                message=frappe.get_traceback(),
                reference_doctype="Email Account",
                reference_name=self.name
            )
            return None

    def get_inbound_mails(self):
        if self._is_graph_api_enabled():
            return self._fetch_from_graph_api()
        return super().get_inbound_mails()

    def _fetch_from_graph_api(self):
        if not self.enable_incoming:
            return []

        mails = []

        try:
            token_doc = self._get_oauth_token()
            if not token_doc:
                frappe.log_error(
                    title="Mail 365: No Token",
                    message=f"Could not get token for {self.name}. Please re-authorize the Connected App.",
                    reference_doctype="Email Account",
                    reference_name=self.name
                )
                return mails

            access_token = token_doc.get_password("access_token")
            messages = self._call_inbox_api(access_token)
            if not messages:
                return mails

            existing_ids = self._get_existing_ids(messages)
            mails = self._process_messages(messages, existing_ids, access_token)

        except Exception:
            frappe.log_error(
                title=f"Mail 365: Error for {self.name}",
                message=frappe.get_traceback(),
                reference_doctype="Email Account",
                reference_name=self.name
            )

        return mails

    def _call_inbox_api(self, access_token):
        headers = {"Authorization": f"Bearer {access_token}"}

        params = {
            "$select": ",".join([
                "id",
                "internetMessageId",
                "conversationId",
                "from",
                "toRecipients",
                "ccRecipients",
                "subject",
                "body",
                "receivedDateTime",
                "isRead",
                "hasAttachments"
            ]),
            "$orderby": "receivedDateTime desc",
            "$top": int(self.initial_sync_count or 100)
        }

        if self.email_sync_option == "UNSEEN":
            params["$filter"] = "isRead eq false"

        response = requests.get(
            GRAPH_INBOX_URL,
            headers=headers,
            params=params,
            timeout=REQUEST_TIMEOUT
        )

        if response.status_code == 401:
            frappe.log_error(
                title="Mail 365: Auth Failed",
                message="Token expired or invalid. Please re-authorize the Connected App.",
                reference_doctype="Email Account",
                reference_name=self.name
            )
            return []

        if response.status_code != 200:
            frappe.log_error(
                title="Mail 365: API Error",
                message=f"Status: {response.status_code}\n{response.text}",
                reference_doctype="Email Account",
                reference_name=self.name
            )
            return []

        return response.json().get("value", [])

    def _get_existing_ids(self, messages):
        message_ids = [
            msg.get("internetMessageId", "").strip("<>")
            for msg in messages
            if msg.get("internetMessageId")
        ]

        if not message_ids:
            return set()

        existing = frappe.get_all(
            "Communication",
            filters={"message_id": ["in", message_ids]},
            pluck="message_id"
        )

        return set(existing)

    def _process_messages(self, messages, existing_ids, access_token):
        """Convert Microsoft 365 messages to InboundMail objects."""
        from frappe_mail_365.overrides.receive import Mail365InboundMail

        mails = []
        append_to = self._get_append_to()

        for msg in messages:
            message_id = msg.get("internetMessageId", "").strip("<>")

            if not message_id or message_id in existing_ids:
                continue

            try:
                raw_email = self._convert_to_email_format(msg, access_token)

                inbound_mail = Mail365InboundMail(
                    raw_email,
                    self,
                    uid=None,
                    seen_status="SEEN" if msg.get("isRead") else "UNSEEN",
                    append_to=append_to
                )

                # Store Graph API IDs for threading and replies
                inbound_mail.graph_message_id = msg.get("id")
                inbound_mail.conversation_id = msg.get("conversationId")

                mails.append(inbound_mail)

            except Exception:
                frappe.log_error(
                    title="Mail 365: Convert Error",
                    message=f"Message: {message_id}\n{frappe.get_traceback()}",
                    reference_doctype="Email Account",
                    reference_name=self.name
                )

        return mails

    def _get_append_to(self):
        """Get append_to doctype for INBOX folder."""
        if self.use_imap and hasattr(self, "imap_folder"):
            for folder in self.imap_folder:
                if folder.folder_name and folder.folder_name.upper() == "INBOX":
                    return folder.append_to
        return None

    def _convert_to_email_format(self, msg, access_token):
        from email.mime.multipart import MIMEMultipart
        from email.mime.text import MIMEText

        has_attachments = msg.get("hasAttachments", False)
        mail = MIMEMultipart("mixed" if has_attachments else "alternative")

        from_data = msg.get("from", {}).get("emailAddress", {})
        if from_data.get("address"):
            mail["From"] = email.utils.formataddr((
                from_data.get("name", ""),
                from_data.get("address")
            ))

        to_list = msg.get("toRecipients", [])
        if to_list:
            to_addrs = []
            for r in to_list:
                addr = r.get("emailAddress", {})
                if addr.get("address"):
                    to_addrs.append(email.utils.formataddr((
                        addr.get("name", ""),
                        addr.get("address")
                    )))
            if to_addrs:
                mail["To"] = ", ".join(to_addrs)

        cc_list = msg.get("ccRecipients", [])
        if cc_list:
            cc_addrs = []
            for r in cc_list:
                addr = r.get("emailAddress", {})
                if addr.get("address"):
                    cc_addrs.append(email.utils.formataddr((
                        addr.get("name", ""),
                        addr.get("address")
                    )))
            if cc_addrs:
                mail["CC"] = ", ".join(cc_addrs)

        mail["Subject"] = msg.get("subject") or "No Subject"

        msg_id = msg.get("internetMessageId", "")
        if msg_id:
            if not msg_id.startswith("<"):
                msg_id = f"<{msg_id}>"
            mail["Message-ID"] = msg_id

        if msg.get("receivedDateTime"):
            mail["Date"] = msg["receivedDateTime"]

        body_data = msg.get("body", {})
        content = body_data.get("content", "")
        content_type = body_data.get("contentType", "HTML").lower()

        if content_type == "html":
            body_part = MIMEText(content, "html", "utf-8")
        else:
            body_part = MIMEText(content, "plain", "utf-8")

        mail.attach(body_part)

        if has_attachments:
            self._add_attachments(mail, msg.get("id"), access_token)

        return mail.as_string()

    def _add_attachments(self, mail, message_id, access_token):
        """Fetch and add attachments from Microsoft 365."""
        import base64
        from email.mime.base import MIMEBase
        from email import encoders

        url = f"{GRAPH_API_BASE}/me/messages/{message_id}/attachments"
        headers = {"Authorization": f"Bearer {access_token}"}

        try:
            response = requests.get(url, headers=headers, timeout=REQUEST_TIMEOUT)
            if response.status_code != 200:
                return

            for att in response.json().get("value", []):
                content_b64 = att.get("contentBytes", "")
                if not content_b64:
                    continue

                content = base64.b64decode(content_b64)
                filename = att.get("name", "attachment")
                ctype = att.get("contentType", "application/octet-stream")

                if "/" in ctype:
                    maintype, subtype = ctype.split("/", 1)
                else:
                    maintype, subtype = "application", "octet-stream"

                mime_att = MIMEBase(maintype, subtype)
                mime_att.set_payload(content)
                encoders.encode_base64(mime_att)
                mime_att.add_header("Content-Disposition", "attachment", filename=filename)

                mail.attach(mime_att)

        except Exception:
            frappe.log_error(
                title="Mail 365: Attachment Error",
                message=frappe.get_traceback(),
                reference_doctype="Email Account",
                reference_name=self.name
            )

    def send_via_graph_api(self, email_data):
        if not self.enable_outgoing:
            frappe.throw(_("Outgoing email not enabled for this account."))

        token_doc = self._get_oauth_token()
        if not token_doc:
            frappe.throw(_("Could not get OAuth token. Please re-authorize the Connected App."))

        access_token = token_doc.get_password("access_token")

        is_reply, graph_message_id = self._is_reply_email(email_data)
        if is_reply and graph_message_id:
            return self._send_reply_email(graph_message_id, email_data, access_token)

        return self._send_new_email(email_data, access_token)

    def _send_new_email(self, email_data, access_token):
        message = {
            "message": {
                "subject": email_data.get("subject") or "No Subject",
                "body": {
                    "contentType": "HTML",
                    "content": email_data.get("message") or ""
                },
                "toRecipients": self._build_recipients(email_data.get("recipients", ""))
            },
            "saveToSentItems": "true"
        }

        if email_data.get("cc"):
            message["message"]["ccRecipients"] = self._build_recipients(email_data["cc"])

        if email_data.get("attachments"):
            attachments_list = self._build_attachments_for_send(email_data["attachments"])
            if attachments_list:
                message["message"]["attachments"] = attachments_list

        headers = {
            "Authorization": f"Bearer {access_token}",
            "Content-Type": "application/json"
        }

        response = requests.post(
            GRAPH_SEND_URL,
            headers=headers,
            json=message,
            timeout=REQUEST_TIMEOUT
        )

        if response.status_code != 202:
            frappe.log_error(
                title="Mail 365: Send Error",
                message=f"Status: {response.status_code}\n{response.text}",
                reference_doctype="Email Account",
                reference_name=self.name
            )
            response.raise_for_status()

        return True

    def _build_recipients(self, recipients_str):
        """Convert recipient string to Graph API format."""
        from frappe.utils import parse_addr

        if not recipients_str:
            return []

        recipients = []
        for recipient in recipients_str.split(","):
            recipient = recipient.strip()
            if recipient:
                name, email_addr = parse_addr(recipient)
                recipients.append({
                    "emailAddress": {
                        "address": email_addr or recipient,
                        "name": name or ""
                    }
                })

        return recipients

    def _build_attachments_for_send(self, attachments_json):
        """Convert Frappe attachments to Graph API format (base64)."""
        import base64

        if not attachments_json:
            return []

        try:
            attachments = frappe.parse_json(attachments_json)
            if not isinstance(attachments, list):
                return []

            graph_attachments = []
            for attachment in attachments:
                if isinstance(attachment, dict) and attachment.get("fid"):
                    file_doc = frappe.get_doc("File", attachment["fid"])
                    file_content = file_doc.get_content()

                    if file_content:
                        if isinstance(file_content, str):
                            file_content = file_content.encode()

                        encoded_content = base64.b64encode(file_content).decode()
                        graph_attachments.append({
                            "@odata.type": "#microsoft.graph.fileAttachment",
                            "name": file_doc.file_name,
                            "contentType": file_doc.file_type or "application/octet-stream",
                            "contentBytes": encoded_content
                        })

            return graph_attachments

        except Exception:
            frappe.log_error(
                title="Mail 365: Build Attachments Error",
                message=frappe.get_traceback(),
                reference_doctype="Email Account",
                reference_name=self.name
            )
            return []

    def _is_reply_email(self, email_data):
        """
        Check if email is a reply by looking at Communication.in_reply_to.

        Flow:
        - email_data.communication = current Communication (COMM-002)
        - COMM-002.in_reply_to = parent Communication (COMM-001)
        - COMM-001.graph_message_id = Graph API message ID for /reply

        Returns (is_reply: bool, graph_message_id: str or None)
        """
        communication_name = email_data.get("communication")
        if not communication_name:
            return False, None

        try:
            comm = frappe.get_doc("Communication", communication_name)

            if not comm.in_reply_to:
                return False, None

            parent = frappe.get_doc("Communication", comm.in_reply_to)
            graph_message_id = getattr(parent, "graph_message_id", None)

            if graph_message_id:
                return True, graph_message_id

        except Exception:
            frappe.log_error(
                title="Mail 365: Reply Detection Error",
                message=frappe.get_traceback(),
                reference_doctype="Email Account",
                reference_name=self.name
            )

        return False, None

    def _send_reply_email(self, graph_message_id, email_data, access_token):
        """
        Send reply using Graph API POST /me/messages/{id}/reply endpoint.

        This maintains email threading in Outlook/Graph API.
        """
        url = f"{GRAPH_API_BASE}/me/messages/{graph_message_id}/reply"
        headers = {
            "Authorization": f"Bearer {access_token}",
            "Content-Type": "application/json"
        }

        payload = {
            "message": {
                "toRecipients": self._build_recipients(email_data.get("recipients", ""))
            },
            "comment": email_data.get("message") or ""
        }

        if email_data.get("cc"):
            payload["message"]["ccRecipients"] = self._build_recipients(email_data["cc"])

        response = requests.post(url, headers=headers, json=payload, timeout=REQUEST_TIMEOUT)

        if response.status_code != 202:
            frappe.log_error(
                title="Mail 365: Reply Error",
                message=f"Status: {response.status_code}\n{response.text}",
                reference_doctype="Email Account",
                reference_name=self.name
            )
            response.raise_for_status()

        return True
