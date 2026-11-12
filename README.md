# Carryon

Onboarding and waitlist for newcomers who want their credit history from home to count in the US. Built by Team 213 (Himali Saini, Tasos Gerasis, Varun Kulkarni, Elena Hu) for TECH 5900 Product Studio at Cornell Tech.

The page follows the Arrive prototype's onboarding screens (Welcome, Verify it's you, Scan your PAN card, Bring your credit history, Unlock your US credit profile, Cards matched to you) and ends on a waitlist. It never collects a passport, visa, selfie or PAN number. It runs two experiments from our de-risking plan:

- **Consent screen test (Card 3).** Each visitor sees either a plain consent screen or one that lists what we read and never touch, with the team's names. We count who gives permission and who asks to share their CIBIL report early.
- **Price test.** Each visitor sees one founding price: $19, $29 or $49. We count who reserves a spot.

It also records when people arrive, what they need credit for first, what has happened to them so far, and how they heard about us.

## How data is saved

`index.html` sends each step to a Google Apps Script web app (`apps-script/Code.gs`), which writes one row per visitor to the **Responses** tab of the "Carryon waitlist" Google Sheet. The **Experiments** tab summarizes the results.

Until `CONFIG.ENDPOINT` in `index.html` holds the web app URL, the page runs in preview mode and saves nothing.

## Testing links

Add `?v=plain` or `?v=controls` to force a consent version, and `?p=19`, `?p=29` or `?p=49` to force a price. `?ref=CODE` records an invite, and `?src=NAME` records a channel (for example `?src=nyu-email`).
