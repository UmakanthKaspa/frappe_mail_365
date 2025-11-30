import frappe
from frappe.core.doctype.communication.email import make as original_make


@frappe.whitelist()
def make(in_reply_to=None, **kwargs):
    """Wrapper that adds in_reply_to support for Graph API replies."""
    result = original_make(**kwargs)

    if in_reply_to and result.get("name"):
        frappe.db.set_value("Communication", result["name"], "in_reply_to", in_reply_to)

    return result
