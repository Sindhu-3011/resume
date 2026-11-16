# Microsoft Teams Interview Invites — Setup Guide (No Azure Experience Needed)

This app can automatically create a Microsoft Teams meeting and send an
Outlook calendar invite whenever someone schedules an interview. This guide
walks you through turning that on, one click at a time — you don't need to
know anything about Azure beforehand.

**Until you finish this guide**, interview scheduling still works exactly as
it does today — nothing breaks. Every invite attempt is just recorded as
"Skipped (not configured)" with no error, until the setup below is complete.

---

## Before you start: do you have what you need?

**You do NOT need to create a new account.** You need admin access to your
**company's existing Microsoft 365 account** — the same one your company
already uses for Outlook email and Teams meetings. Every organization that
uses Teams already has this sitting in the background, whether anyone has
used its admin side before or not.

- **If you're the person who manages your company's IT/Microsoft 365
  settings** (or you have "Global Administrator" access) — you're good to
  go, continue to Step 1.
- **If you're not sure, or you know someone else handles that** — send them
  this whole document. They can do Steps 1–7 in about 10 minutes and just
  hand you 4 short values at the end (Step 8), which is all this app
  actually needs.
- You'll also need **one email inbox** at your company that can act as the
  "host" of every interview meeting — e.g. `interviews@yourcompany.com`, or
  even an existing mailbox like HR's. More on this in Step 7.

---

## Step 1 — Sign in to the Azure Portal

1. Open **[portal.azure.com](https://portal.azure.com)** in your browser.
2. Sign in with your **work email** — the same one you use for Outlook/Teams
   at your company. (Not a personal Gmail or Outlook.com account.)

---

## Step 2 — Open "Microsoft Entra ID"

1. At the very top of the page, there's a search bar. Click it and type:
   **Microsoft Entra ID**
2. Click it in the results.

> This page represents your whole company's Microsoft account setup — all
> your users, security settings, and (what we need) a list of registered
> apps. (You might also see this called **"Azure Active Directory"** in
> older screenshots or documentation — Microsoft renamed it, but it's the
> exact same thing.)

---

## Step 3 — Register this app

This step is just telling Microsoft: "this piece of software is going to
need permission to create meetings on our behalf."

1. In the menu on the left side of the page, find and click
   **"App registrations"**.
2. Click **"+ New registration"** near the top of the page.
3. Fill in the form that appears:
   | Field | What to enter |
   |---|---|
   | **Name** | `TalentEdge AI - Interview Scheduling` (just a label — doesn't affect anything technical, type whatever helps you recognize it later) |
   | **Supported account types** | Leave it on the first option: **"Accounts in this organizational directory only"** |
   | **Redirect URI** | Leave this completely blank |
4. Click the **"Register"** button at the bottom.

You'll land on a new page for this app — keep it open, you'll need it for
the next two steps.

---

## Step 4 — Copy two ID numbers

You're now on the app's **"Overview"** page. You'll see a list of fields.
Find these two — each looks like a long string of letters and numbers with
dashes, like `a1b2c3d4-5678-90ab-cdef-1234567890ab`:

- **Application (client) ID**
- **Directory (tenant) ID**

Click the small copy icon next to each one, and paste both somewhere safe
for now (a notes app, a password manager) — label them clearly so you don't
mix them up later.

---

## Step 5 — Create a "client secret" (a password for the app)

1. On the left-side menu of the same app page, click
   **"Certificates & secrets"**.
2. Click the **"Client secrets"** tab, then **"+ New client secret"**.
3. A small panel appears:
   | Field | What to enter |
   |---|---|
   | **Description** | `TalentEdge AI Graph secret` |
   | **Expires** | Pick 12 or 24 months if you're not sure what your company normally uses |
4. Click **"Add"**.
5. A new row appears with a **"Value"** column (not "Secret ID" — a
   different column next to it). **Copy that Value right now.**

   ⚠️ This is the *only* time this value is ever shown in full. If you
   navigate away before copying it, you can't get it back — you'd just
   delete that row and make a new secret.

Save this value alongside the two IDs from Step 4.

---

## Step 6 — Give the app permission to create meetings and invites

1. On the left-side menu, click **"API permissions"**.
2. Click **"+ Add a permission"**.
3. A panel slides in showing a list of Microsoft services. Click
   **"Microsoft Graph"** (usually the first, biggest tile).
4. You'll be asked to pick between two types of permission — click
   **"Application permissions"** (not "Delegated"). This matters: it means
   the app works on its own, not pretending to be any one specific person.
5. A search box appears. Type and check the box next to:
   - `OnlineMeetings.ReadWrite.All`
6. Clear the search box, type again, and check the box next to:
   - `Calendars.ReadWrite`
7. Click **"Add permissions"** at the bottom of the panel.

You'll now see both permissions listed, each with a **yellow warning
triangle** saying something like "Not granted for [your company]." That's
expected — the next step fixes it.

**Grant consent:** Still on this same page, click the button near the top
that says **"Grant admin consent for [your company name]"**, then confirm
**"Yes"** on the popup. The yellow triangles should turn into **green
checkmarks**.

> If that button is greyed out or you don't see it, your account doesn't
> have admin rights for this last click. Everything you've set up so far is
> saved — just ask someone with Global Administrator access to open this
> same "API permissions" page for this app and click that one button.

---

## Step 7 — Pick the mailbox that "hosts" every meeting

Separately from everything above, pick **one real email inbox** at your
company that will show up as the host/organizer of every interview meeting
this app creates — for example `interviews@yourcompany.com`, or an existing
mailbox like HR's. It just needs to:
- already exist in your company's Microsoft 365, and
- have a Teams/Exchange license (the same kind any normal employee mailbox
  has).

Nothing needs to be installed in this mailbox — it's just used as the
"from" identity when creating meetings.

---

## Step 8 — Hand over the 4 values

By now you should have these 4 things written down:

| # | What it is | Where you got it |
|---|---|---|
| 1 | Directory (tenant) ID | Step 4 |
| 2 | Application (client) ID | Step 4 |
| 3 | Client secret **value** | Step 5 |
| 4 | Organizer mailbox email | Step 7 |

Add them to your `.env` file (the same file your database/email settings
already live in), using these exact names:

```env
AZURE_TENANT_ID=<value #1>
AZURE_CLIENT_ID=<value #2>
AZURE_CLIENT_SECRET=<value #3>
TEAMS_ORGANIZER_UPN=<value #4, e.g. interviews@yourcompany.com>
```

Then **restart the app** so it picks up the new values.

(If someone else did Steps 1–7 for you, just ask them for these same 4
values — nothing else from the Azure Portal is needed.)

---

## Step 9 — Check that it worked

1. Make sure the app was restarted after adding the values above.
2. On the Dashboard, schedule a test interview for any candidate.
3. Within a few seconds, the **"Teams Meeting Link"** column for that row
   should show a **Join** button, instead of "Not Sent" or "Failed". Click
   it — it should open a real Teams meeting.
4. Check the organizer mailbox's calendar (the one from Step 7) — the
   interview should appear there, with the candidate's and interviewer's
   email listed as attendees (if you entered an interviewer email when
   scheduling).
5. **If it shows "Failed" instead:** the most common causes are —
   - The green checkmarks from Step 6 never appeared (admin consent wasn't
     actually granted).
   - The organizer mailbox from Step 7 doesn't exist or has no license.
   - One of the 4 values in Step 8 has a typo or extra space.

---

## Keeping it working: rotating the secret

The "client secret" you created in Step 5 has an expiry date you chose.
Before it expires, repeat Step 5 to create a new one, update
`AZURE_CLIENT_SECRET` in your `.env` file with the new value, and restart
the app. Once that's confirmed working, you can delete the old secret from
the **"Certificates & secrets"** page.
