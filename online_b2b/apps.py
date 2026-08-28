from django.apps import AppConfig


class OnlineB2BConfig(AppConfig):
    name = "online_b2b"
    # Admin-panel section header. The app is *code-named* online_b2b for legacy
    # reasons, but its models are the whole ORDER STORE (online B2B **and**
    # offline) + shared master data — so the admin shows a segment-neutral name.
    # Display-only: does NOT change the app_label / imports / URLs.
    verbose_name = 'Orders & Master Data'
