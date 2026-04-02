# Outlook Email Evaluator — Android/Mobile Add-in

AI-powered phishing and spam detection for **Outlook on Android** (and iOS).

Same Supabase backend as the Chrome extension and Desktop add-in.

## Files

| File | Purpose |
|---|---|
| `manifest.xml` | Office Add-in manifest with `MobileFormFactor` extension point |
| `taskpane.html` | Task pane shell — mobile-first layout |
| `taskpane.css` | CSS optimised for narrow (~360px) task panes, large touch targets |
| `taskpane.js` | All logic: Office.js email reading, Supabase fetch, result rendering |
| `commands.html` | Required stub for the `FunctionFile` manifest reference |

## How it works

Unlike the Chrome extension (which scrapes the Outlook Web DOM), this add-in uses the
**Office.js Mailbox API** to read the email — the same approach as the desktop add-in.
This means it works natively on Android without any browser tricks.

The Supabase edge function (`analyze-email`) is **shared** with the Chrome and desktop add-ins.
No backend changes needed.

## Deploying

1. Host these files on GitHub Pages under `a1ch.github.io/outlook-email-evaluator-android/`
2. Create an `icons/` folder with icon-16.png, icon-32.png, icon-64.png, icon-80.png, icon-128.png
   (copy from the desktop add-in repo)
3. Sideload the manifest via the Microsoft 365 Admin Center or Outlook Web:
   - Outlook Web → Settings → Manage add-ins → Add from file → upload `manifest.xml`
4. On Android, open the Outlook app → open any email → tap the `...` menu → **Email Evaluator**

## Requirements

- Microsoft 365 account (work/school)
- Your M365 tenant must allow custom/sideloaded add-ins
- The add-in appears on both Android and iOS automatically once sideloaded

## Settings (configured in the add-in ⚙️ menu)

| Setting | Value |
|---|---|
| Supabase Proxy URL | `https://pikplhvawbhndijpkdbq.supabase.co/functions/v1/analyze-email` |
| Extension Token | Same token used in Chrome and desktop add-ins |
| Organization Domain | e.g. `yourcompany.com` — flags external senders |
| Additional Instructions | Extra context for the AI |

## Differences from the desktop add-in

- No dark mode toggle (Android respects system dark mode via CSS `prefers-color-scheme`)
- Settings accessed via ⚙️ button → full-screen settings panel (better for mobile)
- Larger touch targets throughout (44px+ tap areas)
- Sender shown in email header for quick glance
- No DOM manipulation tricks — pure Office.js API
