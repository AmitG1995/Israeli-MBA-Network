# Sign-up backend (Google Sheet + Gmail)

`Code.gs` is a small Google Apps Script that handles real sign-ups:

1. A new member submits the sign-up form. Their request is saved as a row in a Google Sheet.
2. The admin gets an email with the applicant's details and a **Review and approve** button.
3. The button opens a page where the admin clicks **Approve** or **Decline**. Approving marks the member approved in the Sheet and sends them the official welcome email.
4. The next time the member opens the app, it sees they've been approved.

Emails are sent from the Google account that sets this up. Free Gmail accounts can send about 100 emails a day.

## One-time setup (about 10 minutes)

1. **Create the Sheet.** Go to [sheets.new](https://sheets.new) while signed in to the admin's Google account, and name it "Israeli MBA Network – Sign-ups".
2. **Add the script.** In the Sheet, open **Extensions → Apps Script**. Delete the sample code, paste in all of `Code.gs` from this folder, and click **Save**.
3. **Set the admin email (optional).** Approval emails go to the Google account that runs the script. To send them somewhere else, open **Project Settings** (gear icon) → **Script properties** → **Add script property**, with the name `ADMIN_EMAIL` and your email as the value.
4. **Deploy it.** Click **Deploy → New deployment**. Under the gear next to "Select type", choose **Web app**, then set:
   - **Execute as:** Me
   - **Who has access:** Anyone

   Click **Deploy**, then **Authorize access**. Google asks for permission to manage the Sheet and send email as you. If it says "Google hasn't verified this app", click **Advanced → Go to (project name)**. That's expected for your own script.
5. **Copy the Web app URL.** It ends in `/exec`.
6. **Connect the app.** In `index.html`, find `const BACKEND_URL='';` and paste the URL between the quotes. Commit that change on a feature branch, like any other.

## Log in with LinkedIn (one-time setup)

Members log in with LinkedIn, and the backend lets them in if their LinkedIn email matches an approved request. Members can also get a login code by email instead.

1. **Create a LinkedIn app.** Go to [linkedin.com/developers/apps](https://www.linkedin.com/developers/apps) → **Create app**. LinkedIn requires a LinkedIn Page for the app. If you don't have one, the form links to creating one.
2. **Turn on sign-in.** On the app's **Products** tab, request **Sign In with LinkedIn using OpenID Connect**. Access is granted right away.
3. **Add redirect URLs.** On the **Auth** tab, under **Authorized redirect URLs**, add every address the app runs at, exactly as it appears in the browser, e.g. `http://localhost:8010/index.html`. Add the hosted address too once the site is live.
4. **Give the backend the keys.** From the **Auth** tab, copy the **Client ID** and **Primary Client Secret**. In Apps Script, open **Project Settings → Script properties** and add `LINKEDIN_CLIENT_ID` and `LINKEDIN_CLIENT_SECRET`. The secret must never go in `index.html` or GitHub.
5. **Give the app the Client ID.** In `index.html`, set `const LINKEDIN_CLIENT_ID='...';`. The Client ID is public, so this is safe.

## Updating the script

After editing `Code.gs`, paste it into Apps Script again, then click **Deploy → Manage deployments → Edit (pencil) → Version: New version → Deploy**. This keeps the same URL.

## Notes

- **The Sheet is your member list.** Its `status` column says `pending`, `approved` or `declined`. Don't rename the columns.
- **Approve links are secret.** Each email has its own private link. Don't forward approval emails.
- **The web app URL is public** once it's in `index.html`, so anyone could submit a request. That's fine, since nobody gets in without the admin approving.
- **Declining doesn't email the applicant.**
