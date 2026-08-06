# Web Portal for Church Ministries 

A web portal built for church ministries to centralise announcements, resources and information for Co-workers (i.e. the Co-workers handbook).

This was built for Church of Singapore's Music Ministry but any other church ministry is welcome to use this and adapt this as needed for your ministry. 

🔗 **[View the live site here](https://cosmusicministry.onrender.com)** 

---

## Contents of Web Portal

**Pages**
- **Home** — Contains a hero section, a welcome message and navigation panel. 
- **Announcements** — Latest announcements from the ministry, with photos and tags
- **Handbook** — The Co-workers Handbook, organised by section and chapter, with a table of contents for easy navigation
- **Resources** — Articles and reference material for Co-workers

**Features**
- Hero section on home page and a welcome message to welcome Co-workers and visitors of the site
- Smooth page animations and mobile-friendly layout
- Handbook is searchable and easy to navigate across sections and chapters
- Announcements are tagged by category (e.g. General, Events, Worship)

---

## For Admins

Admins can manage all site content through the built-in admin portal.

**Accessing the portal**

Go to `/admin` on the live site and enter the admin password.

**What you can do**
- **Announcements** — Add, edit, or remove announcements. You can upload an image, set a date, and tag each post.
- **Handbook** — Add or edit sections and chapters. The editor supports rich text formatting (bold, italics, headings, lists).
- **Export** — Download the full Co-workers Handbook as a Word document (.docx) at any time.

---

## For Developers

This repository contains the frontend templates and static assets. Sensitive configuration (database credentials, API keys) is excluded.

**Tech stack**
- **Backend:** Python / Flask
- **Database:** PostgreSQL (Supabase)
- **Image hosting:** Cloudinary
- **Deployment:** Render

**To run locally**

1. Clone the repo and install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

2. Fill in the `.env` file with the following:
   ```
   DATABASE_URL=your_supabase_connection_string
   SECRET_KEY=your_secret_key
   ADMIN_PASSWORD=your_admin_password
   CLOUDINARY_CLOUD_NAME=your_cloud_name
   CLOUDINARY_API_KEY=your_api_key
   CLOUDINARY_API_SECRET=your_api_secret
   ```

3. Run the app:
   ```bash
   python app.py
   ```

The app will be available at `http://localhost:5003`.
