from django.contrib.auth.models import User
from django.test import Client, TestCase
from django.urls import reverse


class CoreViewsTestCase(TestCase):
    def setUp(self):
        self.client = Client()
        self.user = User.objects.create_user(username="testuser", password="testpassword")

    def test_home_view(self):
        response = self.client.get(reverse("home"))
        self.assertEqual(response.status_code, 200)
        self.assertTemplateUsed(response, "core/landing.html")

    def test_departments_view(self):
        # login first
        self.client.login(username="testuser", password="testpassword")
        response = self.client.get(reverse("departments"))
        self.assertEqual(response.status_code, 200)
        self.assertTemplateUsed(response, "core/departments.html")

    def test_departments_view_redirects_if_not_logged_in(self):
        response = self.client.get(reverse("departments"))
        self.assertEqual(response.status_code, 302)
        self.assertTrue(response.url.startswith(reverse("login")))
