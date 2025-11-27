import frappe
import click


def before_uninstall():
    remove_email_account_field()
    remove_communication_fields()
    frappe.db.commit()


def remove_email_account_field():
    meta = frappe.get_meta("Email Account")

    if meta.has_field("use_graph_api"):
        click.secho("* Removing Custom Field use_graph_api from Email Account", fg="red")

        custom_field_name = "Email Account-use_graph_api"
        if frappe.db.exists("Custom Field", {"name": custom_field_name}):
            frappe.delete_doc("Custom Field", custom_field_name, force=1)

    frappe.clear_cache(doctype="Email Account")


def remove_communication_fields():
    meta = frappe.get_meta("Communication")

    if meta.has_field("graph_message_id"):
        click.secho("* Removing Custom Field graph_message_id from Communication", fg="red")

        custom_field_name = "Communication-graph_message_id"
        if frappe.db.exists("Custom Field", {"name": custom_field_name}):
            frappe.delete_doc("Custom Field", custom_field_name, force=1)

    if meta.has_field("conversation_id"):
        click.secho("* Removing Custom Field conversation_id from Communication", fg="red")

        custom_field_name = "Communication-conversation_id"
        if frappe.db.exists("Custom Field", {"name": custom_field_name}):
            frappe.delete_doc("Custom Field", custom_field_name, force=1)

    frappe.clear_cache(doctype="Communication")
