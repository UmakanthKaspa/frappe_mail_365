import frappe
from frappe.email.doctype.email_queue.email_queue import EmailQueue


class Mail365EmailQueue(EmailQueue):
    """Extended Email Queue with Graph API sending support."""

    def send(self, smtp_server_instance=None, frappe_mail_client=None, force_send=False):
        """Send emails via Graph API or SMTP."""
        if not self.can_send_now() and not force_send:
            return

        email_account_doc = self.get_email_account(raise_error=True)

        if getattr(email_account_doc, "use_graph_api", 0) and email_account_doc.enable_outgoing:
            return self._send_via_graph_api(email_account_doc)

        return super().send(smtp_server_instance, frappe_mail_client, force_send)

    def _send_via_graph_api(self, email_account_doc):
        """Send email using Microsoft Graph API instead of SMTP."""
        try:
            self.update_status(status="Sending", commit=True)

            from email.parser import Parser
            from email.policy import SMTP as SMTPPolicy

            message_obj = Parser(policy=SMTPPolicy).parsestr(self.message)
            subject = message_obj["Subject"] or "No Subject"
            cc_str = self.show_as_cc or ""
            cc_list = [e.strip() for e in cc_str.split(",") if e.strip()] if cc_str else []
            to_recipients = [r.recipient for r in self.recipients if r.recipient not in cc_list]
            body = self._extract_body(message_obj)

            email_data = frappe._dict({
                "subject": subject,
                "recipients": ", ".join(to_recipients),
                "cc": cc_str if cc_str else None,
                "message": body or self.message,
                "attachments": self.attachments,
                "communication": self.communication,
            })

            email_account_doc.send_via_graph_api(email_data)

            for recipient in self.recipients:
                if not recipient.is_mail_sent():
                    recipient.update_db(status="Sent", commit=True)

            self.update_status(status="Sent", commit=True)

        except Exception:
            frappe.log_error(
                title=f"Mail 365: Send Error for {self.name}",
                message=frappe.get_traceback(),
                reference_doctype="Email Queue",
                reference_name=self.name
            )
            self.update_status("Error", commit=True)
            raise

    def _extract_body(self, message_obj):
        """Extract HTML or plain text body from email message."""
        body = None

        if message_obj.is_multipart():
            for part in message_obj.walk():
                content_type = part.get_content_type()
                if content_type == "text/html":
                    body = part.get_content()
                    break
                elif content_type == "text/plain" and not body:
                    body = part.get_content()
        else:
            body = message_obj.get_content()

        return body
