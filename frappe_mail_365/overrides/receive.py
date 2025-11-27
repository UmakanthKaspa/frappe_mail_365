from frappe.email.receive import InboundMail as BaseInboundMail


class Mail365InboundMail(BaseInboundMail):
    """Extended InboundMail with Microsoft 365 conversation threading."""

    def parent_communication(self):
        """Find parent email using conversationId when In-Reply-To fails."""
        from frappe.core.doctype.communication.communication import Communication

        parent = super().parent_communication()
        if parent:
            return parent

        if hasattr(self, "conversation_id") and self.conversation_id:
            parent = Communication.find_one_by_filters(
                conversation_id=self.conversation_id,
                email_account=self.email_account.name,
                creation=("<", self.date),
                order_by="creation DESC"
            )

            if parent:
                self._parent_communication = parent
                return parent

        return ""

    def reference_document(self):
        """Find reference document using conversation's first email."""
        from frappe.core.doctype.communication.communication import Communication

        ref_doc = super().reference_document()
        if ref_doc:
            return ref_doc

        if hasattr(self, "conversation_id") and self.conversation_id:
            first_email = Communication.find_one_by_filters(
                conversation_id=self.conversation_id,
                email_account=self.email_account.name,
                order_by="creation ASC"
            )

            if first_email and first_email.reference_doctype and first_email.reference_name:
                ref_doc = self.get_doc(
                    first_email.reference_doctype,
                    first_email.reference_name,
                    ignore_error=True
                )

                if ref_doc:
                    self._reference_document = ref_doc
                    return ref_doc

        return ""

    def as_dict(self):
        """Add Microsoft 365 fields to Communication data."""
        data = super().as_dict()

        if hasattr(self, "graph_message_id") and self.graph_message_id:
            data["graph_message_id"] = self.graph_message_id

        if hasattr(self, "conversation_id") and self.conversation_id:
            data["conversation_id"] = self.conversation_id

        return data
