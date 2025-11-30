/**
 * Add in_reply_to support for Microsoft 365 Graph API reply threading.
 *
 * This patches CommunicationComposer.send_email to inject parent Communication
 * name when replying. The global frappe.call patch is safe because:
 * 1. Only applied when in_reply_to exists (replying to an email)
 * 2. Only modifies calls to "frappe.core.doctype.communication.email.make"
 * 3. All other frappe.call invocations pass through unchanged
 * 4. Original is restored immediately after send_email returns
 */
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
