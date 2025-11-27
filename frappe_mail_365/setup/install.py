import frappe
import click
from frappe.custom.doctype.custom_field.custom_field import create_custom_fields


def after_install():
    add_email_account_field()
    add_communication_fields()


def add_email_account_field():
    meta = frappe.get_meta("Email Account")

    if meta.has_field("use_graph_api"):
        click.secho("  - use_graph_api already exists", fg="yellow")
        return

    click.secho("  - Adding 'Use Microsoft 365 Graph API' field", fg="cyan")

    create_custom_fields({
        "Email Account": [
            {
                "fieldname": "use_graph_api",
                "fieldtype": "Check",
                "label": "Use Microsoft 365 Graph API",
                "default": "0",
                "insert_after": "service",
                "description": "Use Graph API instead of SMTP/IMAP for Microsoft 365 accounts.",
            },
        ]
    })

    frappe.clear_cache(doctype="Email Account")


def add_communication_fields():
    meta = frappe.get_meta("Communication")

    if not meta.has_field("graph_message_id"):
        click.secho("  - Adding 'graph_message_id' field", fg="cyan")
        create_custom_fields({
            "Communication": [
                {
                    "fieldname": "graph_message_id",
                    "fieldtype": "Small Text",
                    "label": "Graph Message ID",
                    "insert_after": "uid",
                    "read_only": 1,
                },
            ]
        })

    if not meta.has_field("conversation_id"):
        click.secho("  - Adding 'conversation_id' field", fg="cyan")
        create_custom_fields({
            "Communication": [
                {
                    "fieldname": "conversation_id",
                    "fieldtype": "Small Text",
                    "label": "Conversation ID",
                    "insert_after": "graph_message_id",
                    "read_only": 1,
                },
            ]
        })

    frappe.clear_cache(doctype="Communication")