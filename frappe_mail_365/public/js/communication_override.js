frappe.ready(function() {
    if (!frappe.views?.CommunicationComposer) return;

    const original_send_email = frappe.views.CommunicationComposer.prototype.send_email;

    frappe.views.CommunicationComposer.prototype.send_email = function(btn, form_values, selected_attachments, print_html, print_format) {
        const me = this;
        const in_reply_to = me.last_email?.name;

        if (in_reply_to) {
            const original_call = frappe.call;
            frappe.call = function(opts) {
                if (opts.method === "frappe.core.doctype.communication.email.make") {
                    opts.args.in_reply_to = in_reply_to;
                }
                return original_call.call(this, opts);
            };

            const result = original_send_email.call(this, btn, form_values, selected_attachments, print_html, print_format);
            frappe.call = original_call;
            return result;
        }

        return original_send_email.call(this, btn, form_values, selected_attachments, print_html, print_format);
    };
});
