# RENEE Warehouse Documentation

## Overview

RENEE Warehouse is a modular Django application designed to manage various operational departments of RENEE Cosmetics. It centralizes common functionality (authentication, dashboard navigation) within a `core` app while delegating department-specific tasks to dedicated apps (e.g., `offline`, `inventory`, `returns`).

## System Architecture

*   **Django Backend:** The project heavily utilizes Django's Class-Based Views (CBVs) for standard web routing and form handling.
*   **Modular Applications:**
    *   `core`: Contains the global login logic, application routing, UI layouts, styling, and user profile management.
    *   `offline`: Contains functionality for offline retail, currently featuring the GT Mass Dump tool which uses `pandas` and `openpyxl` to process Excel files.
    *   *Placeholders:* `online_b2b`, `online_b2c`, `returns`, `inventory`, `other` exist to route future functionalities.
*   **Authentication:** Access to any page besides the login portal requires an authenticated user. Django's `LoginRequiredMixin` is applied to all internal views.

## User Interface (UI) Components

The application follows a clean, black-and-white minimalist design language inspired by RENEE Cosmetics branding.

### The Header
*   A persistent element defined in `core/templates/core/base.html`.
*   **Left Side:** Contains the RENEE logo and dynamic breadcrumb navigation (`{% block breadcrumb %}`).
*   **Right Side:** Contains the User Account menu (linking to the profile editing page) and a "Log Out" button.

### Toasts / Notifications
*   Used for global system messaging (Success, Error, Warning, Info).
*   Messages are rendered via Django's `messages` framework.
*   **Behavior:** Toasts slide in from the top right, remain visible for exactly 3 seconds, and then trigger a slow CSS fade-out animation before being removed from the DOM.
*   *Note:* The UI size of toasts is kept minimal and unobtrusive to ensure a clean interface.

### Forms
*   All forms (Login, Profile Editing, Password Change) utilize the unified `.warehouse-login` / `.warehouse-form-container` CSS classes for consistent styling. Form inputs feature a clean 1px border that turns black on focus.

## Flow & Navigation

1.  **Authentication:** Users land on the `HomeView` (root `/`) to log in. Upon a successful login, a green success toast is displayed, and the user is redirected to the Departments Dashboard.
    *   **Account Creation:** Users can click "Create one here" to navigate to the `SignUpView` (`/signup/`). This form uses the standard Django `UserCreationForm`. Upon success, users are automatically logged in and redirected.
    *   **Logout:** Users can click "Log Out" in the header. The `CustomLogoutView` processes this and triggers a "Logout successful" toast message.
2.  **Departments Dashboard:** Displays a grid of available applications (`DepartmentsView`). Active departments (e.g., `Offline`) are clickable, while inactive ones display a "Coming Soon" label.
3.  **Department Applications:** Clicking a department routes the user to that department's specific dashboard (e.g., `offline/dashboard.html`), from which individual tools can be accessed.
4.  **My Account:** Users can click their username in the header to access the `ProfileView`. This page allows users to update their First Name, Last Name, and Email Address. It also provides a link to the `CustomPasswordChangeView` for secure password modifications.

## Frontend Standardization

*   **CSS Architecture:** All styles are centralized in `core/static/core/css/style.css`. Inline styles (`style="..."`) and inline scripts (`<script>`) are strictly avoided in templates to ensure a clean Content Security Policy (CSP) and maintainable codebase.
*   **Forms:** Forms utilize semantic HTML (e.g., `<label>`, `<input>`) and are wrapped in the `.warehouse-form-container` class. Help text and error lists (`.help-text`, `.errorlist`) are styled natively to fit the RENEE branding without relying on default browser lists or bullet points.