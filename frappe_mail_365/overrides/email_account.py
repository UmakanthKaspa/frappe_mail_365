import email.utils
import requests
import frappe
from frappe import _
from frappe.email.doctype.email_account.email_account import EmailAccount


GRAPH_API_BASE = "https://graph.microsoft.com/v1.0"
GRAPH_INBOX_URL = f"{GRAPH_API_BASE}/me/mailFolders/inbox/messages"
REQUEST_TIMEOUT = 30


class Mail365EmailAccount(EmailAccount):
    """Email Account with Microsoft 365 Graph API support."""

    def validate(self):
        """Validate Graph API settings before saving."""
        if self._is_graph_api_enabled():
            self._validate_graph_settings()
        super().validate()

    def validate_smtp_conn(self):
        """Skip SMTP validation for Graph API."""
        if self._is_graph_api_enabled():
            return None
        return super().validate_smtp_conn()

    def validate_imap(self):
        """Skip IMAP validation for Graph API."""
        if self._is_graph_api_enabled():
            return None
        return super().validate_imap()

    def get_incoming_server(self, *args, **kwargs):
        """Skip IMAP server setup for Graph API."""
        if self._is_graph_api_enabled():
            return None
        return super().get_incoming_server(*args, **kwargs)

    def _is_graph_api_enabled(self):
        """Check if Graph API is enabled."""
        return bool(getattr(self, "use_graph_api", 0))

    def _validate_graph_settings(self):
        """Validate OAuth configuration for Graph API."""
        if self.auth_method != "OAuth":
            frappe.throw(_("Microsoft 365 Graph API requires OAuth authentication."))

        if not self.connected_app:
            frappe.throw(_("Please select a Connected App for Microsoft 365."))

    def _get_oauth_token(self):
        """Get valid OAuth access token from Connected App."""
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
        """Override to use Graph API instead of IMAP."""
        if self._is_graph_api_enabled():
            return self._fetch_from_graph_api()
        return super().get_inbound_mails()

    def _fetch_from_graph_api(self):
        """Fetch emails from Microsoft 365 Graph API."""
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
        """Call Microsoft 365 Graph API to get inbox emails."""
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
        """Get message IDs that already exist in Communication."""
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

    # TODO: Works only for INBOX for now, need to improve for all folders
    def _get_append_to(self):
        """Get doctype to link emails to."""
        if self.use_imap and hasattr(self, "imap_folder"):
            for folder in self.imap_folder:
                if folder.folder_name and folder.folder_name.upper() == "INBOX":
                    return folder.append_to

    def _convert_to_email_format(self, msg, access_token):
        """Convert Microsoft 365 JSON to RFC822 email format."""
        from email.mime.multipart import MIMEMultipart
        from email.mime.text import MIMEText

        has_attachments = msg.get("hasAttachments", False)
        mail = MIMEMultipart("mixed" if has_attachments else "alternative")

        # From
        from_data = msg.get("from", {}).get("emailAddress", {})
        if from_data.get("address"):
            mail["From"] = email.utils.formataddr((
                from_data.get("name", ""),
                from_data.get("address")
            ))

        # To
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

        # CC
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

        # Subject
        mail["Subject"] = msg.get("subject") or "No Subject"

        # Message-ID
        msg_id = msg.get("internetMessageId", "")
        if msg_id:
            if not msg_id.startswith("<"):
                msg_id = f"<{msg_id}>"
            mail["Message-ID"] = msg_id

        # Date
        if msg.get("receivedDateTime"):
            mail["Date"] = msg["receivedDateTime"]

        # Body
        body_data = msg.get("body", {})
        content = body_data.get("content", "")
        content_type = body_data.get("contentType", "HTML").lower()

        if content_type == "html":
            body_part = MIMEText(content, "html", "utf-8")
        else:
            body_part = MIMEText(content, "plain", "utf-8")

        mail.attach(body_part)

        # Attachments
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
