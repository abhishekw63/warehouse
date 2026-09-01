from django.contrib.auth.models import User
from django.test import Client, TestCase
from django.urls import reverse


class CoreViewsTestCase(TestCase):
    def setUp(self):
        self.client = Client()
        self.user = User.objects.create_user(username='testuser', password='testpassword')

    def test_home_view(self):
        response = self.client.get(reverse('home'))
        self.assertEqual(response.status_code, 200)
        self.assertTemplateUsed(response, 'core/landing.html')

    def test_departments_redirects_to_order_management(self):
        # The departments picker was removed — the URL name is kept so the many
        # breadcrumb {% url 'departments' %} refs still resolve, and it now sends
        # the user straight to the Order Management workspace.
        self.client.login(username='testuser', password='testpassword')
        response = self.client.get(reverse('departments'))
        self.assertEqual(response.status_code, 302)
        self.assertEqual(response.url, reverse('b2b_dashboard'))