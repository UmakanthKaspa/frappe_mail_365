app_name = "frappe_mail_365"
app_title = "Frappe Mail 365"
app_publisher = "Umakanth Kaspa"
app_description = "Microsoft 365 email integration for Frappe using Graph API"
app_email = "kaspaumakanth1999@gmail.com"
app_license = "mit"

# Apps
# ------------------

# required_apps = []

# Each item in the list will be shown as an app in the apps page
# add_to_apps_screen = [
# 	{
# 		"name": "frappe_mail_365",
# 		"logo": "/assets/frappe_mail_365/logo.png",
# 		"title": "Frappe Mail 365",
# 		"route": "/frappe_mail_365",
# 		"has_permission": "frappe_mail_365.api.permission.has_app_permission"
# 	}
# ]

# Includes in <head>
# ------------------

# include js, css files in header of desk.html
# app_include_css = "/assets/frappe_mail_365/css/frappe_mail_365.css"
app_include_js = "/assets/frappe_mail_365/js/communication_override.js"

# include js, css files in header of web template
# web_include_css = "/assets/frappe_mail_365/css/frappe_mail_365.css"
# web_include_js = "/assets/frappe_mail_365/js/frappe_mail_365.js"

# include custom scss in every website theme (without file extension ".scss")
# website_theme_scss = "frappe_mail_365/public/scss/website"

# include js, css files in header of web form
# webform_include_js = {"doctype": "public/js/doctype.js"}
# webform_include_css = {"doctype": "public/css/doctype.css"}

# include js in page
# page_js = {"page" : "public/js/file.js"}

# include js in doctype views
# doctype_js = {"doctype" : "public/js/doctype.js"}
# doctype_list_js = {"doctype" : "public/js/doctype_list.js"}
# doctype_tree_js = {"doctype" : "public/js/doctype_tree.js"}
# doctype_calendar_js = {"doctype" : "public/js/doctype_calendar.js"}

# Svg Icons
# ------------------
# include app icons in desk
# app_include_icons = "frappe_mail_365/public/icons.svg"

# Home Pages
# ----------

# application home page (will override Website Settings)
# home_page = "login"

# website user home page (by Role)
# role_home_page = {
# 	"Role": "home_page"
# }

# Generators
# ----------

# automatically create page for each record of this doctype
# website_generators = ["Web Page"]

# automatically load and sync documents of this doctype from downstream apps
# importable_doctypes = [doctype_1]

# Jinja
# ----------

# add methods and filters to jinja environment
# jinja = {
# 	"methods": "frappe_mail_365.utils.jinja_methods",
# 	"filters": "frappe_mail_365.utils.jinja_filters"
# }

# Installation
# ------------

# before_install = "frappe_mail_365.install.before_install"
after_install = "frappe_mail_365.setup.install.after_install"

# Uninstallation
# ------------

before_uninstall = "frappe_mail_365.setup.uninstall.before_uninstall"
# after_uninstall = "frappe_mail_365.uninstall.after_uninstall"

# Integration Setup
# ------------------
# To set up dependencies/integrations with other apps
# Name of the app being installed is passed as an argument

# before_app_install = "frappe_mail_365.utils.before_app_install"
# after_app_install = "frappe_mail_365.utils.after_app_install"

# Integration Cleanup
# -------------------
# To clean up dependencies/integrations with other apps
# Name of the app being uninstalled is passed as an argument

# before_app_uninstall = "frappe_mail_365.utils.before_app_uninstall"
# after_app_uninstall = "frappe_mail_365.utils.after_app_uninstall"

# Desk Notifications
# ------------------
# See frappe.core.notifications.get_notification_config

# notification_config = "frappe_mail_365.notifications.get_notification_config"

# Permissions
# -----------
# Permissions evaluated in scripted ways

# permission_query_conditions = {
# 	"Event": "frappe.desk.doctype.event.event.get_permission_query_conditions",
# }
#
# has_permission = {
# 	"Event": "frappe.desk.doctype.event.event.has_permission",
# }

# Document Events
# ---------------
# Hook on document methods and events

# doc_events = {
# 	"*": {
# 		"on_update": "method",
# 		"on_cancel": "method",
# 		"on_trash": "method"
# 	}
# }

# Scheduled Tasks
# ---------------

# scheduler_events = {
# 	"all": [
# 		"frappe_mail_365.tasks.all"
# 	],
# 	"daily": [
# 		"frappe_mail_365.tasks.daily"
# 	],
# 	"hourly": [
# 		"frappe_mail_365.tasks.hourly"
# 	],
# 	"weekly": [
# 		"frappe_mail_365.tasks.weekly"
# 	],
# 	"monthly": [
# 		"frappe_mail_365.tasks.monthly"
# 	],
# }

# Testing
# -------

# before_tests = "frappe_mail_365.install.before_tests"

# Extend DocType Class
# ------------------------------
#
# Specify custom mixins to extend the standard doctype controller.
extend_doctype_class = {
    "Email Account": "frappe_mail_365.overrides.email_account.Mail365EmailAccount",
    "Email Queue": "frappe_mail_365.overrides.email_queue.Mail365EmailQueue",
}

# Overriding Methods
# ------------------------------
override_whitelisted_methods = {
    "frappe.core.doctype.communication.email.make": "frappe_mail_365.overrides.communication.make"
}
#
# each overriding function accepts a `data` argument;
# generated from the base implementation of the doctype dashboard,
# along with any modifications made in other Frappe apps
# override_doctype_dashboards = {
# 	"Task": "frappe_mail_365.task.get_dashboard_data"
# }

# exempt linked doctypes from being automatically cancelled
#
# auto_cancel_exempted_doctypes = ["Auto Repeat"]

# Ignore links to specified DocTypes when deleting documents
# -----------------------------------------------------------

# ignore_links_on_delete = ["Communication", "ToDo"]

# Request Events
# ----------------
# before_request = ["frappe_mail_365.utils.before_request"]
# after_request = ["frappe_mail_365.utils.after_request"]

# Job Events
# ----------
# before_job = ["frappe_mail_365.utils.before_job"]
# after_job = ["frappe_mail_365.utils.after_job"]

# User Data Protection
# --------------------

# user_data_fields = [
# 	{
# 		"doctype": "{doctype_1}",
# 		"filter_by": "{filter_by}",
# 		"redact_fields": ["{field_1}", "{field_2}"],
# 		"partial": 1,
# 	},
# 	{
# 		"doctype": "{doctype_2}",
# 		"filter_by": "{filter_by}",
# 		"partial": 1,
# 	},
# 	{
# 		"doctype": "{doctype_3}",
# 		"strict": False,
# 	},
# 	{
# 		"doctype": "{doctype_4}"
# 	}
# ]

# Authentication and authorization
# --------------------------------

# auth_hooks = [
# 	"frappe_mail_365.auth.validate"
# ]

# Automatically update python controller files with type annotations for this app.
# export_python_type_annotations = True

# default_log_clearing_doctypes = {
# 	"Logging DocType Name": 30  # days to retain logs
# }

# Translation
# ------------
# List of apps whose translatable strings should be excluded from this app's translations.
# ignore_translatable_strings_from = []

