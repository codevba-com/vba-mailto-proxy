# VBA MailToProxy - the CreateEmail procedure
This module contains the CreateEmail procedure which allow you to create an email in the default mail client in edit mode using MailTo. It provides fallbacks to Gmail/Outlook.com. 
- The created email is simple, it only supports plain text.
- There is no way to actually send the mail without the user being involved.
- CC and BCC field implementation have been commented out but you can easily reintroduce them if needed.
- No error handling, any errors occurring will trickle to the calling procedure and should be handled there. 
Calling the procedure is simple:

```vba
CreateEmail "john@doe.com", "about MailTo",  "This is only for no frills emails"
```

## Troubleshooting

The `CreateEmail` procedure relies on `ShellExecute` to launch the default mail client using a `mailto:` URL. If the email window does not appear, one of the following issues is likely the cause:

### **No default mail client is configured**
Windows must have an application associated with the `mailto:` protocol.  
If none is set, `ShellExecute` returns error **31 (SE_ERR_NOASSOC)** and nothing opens.

### **Mail client is installed but not properly registered**
Some email programs fail to register the `mailto:` handler correctly.  
Symptoms include:
- No email window appearing  
- Error **31**  
- The wrong application opening

### **URL too long or contains invalid characters**
`mailto:` URLs have practical length limits. Unencoded characters (spaces, `%`, `&`, line breaks) or very long bodies may cause:
- Silent failure  
- Truncated fields  
- Generic errors such as **2 (SE_ERR_FNF)** or **5 (SE_ERR_ACCESSDENIED)**

### **Mail client rejects or mishandles the request**
Certain clients—especially Outlook desktop—may:
- Ignore parameters  
- Fail to open if another instance is busy  
- Open a blank message instead of the expected content

### **Corporate or system policy restrictions**
In managed environments, protocol handlers like `mailto:` may be disabled.  
This typically results in **SE_ERR_ACCESSDENIED (5)** or no visible action.

### **Webmail handlers misconfigured**
If the user has set Gmail or Outlook.com as the default handler via a browser:
- The browser may fail to launch  
- Parameters may not be passed correctly  
- The wrong profile may open  

The built‑in fallback to Gmail/Outlook.com helps mitigate this scenario.

## How to Fix

If `CreateEmail` fails to open a new message window, try the following steps:

### **1. Verify the default mail client**
Check **Settings → Apps → Default apps → Choose defaults by link type → mailto**  
Ensure a valid desktop mail client (Outlook, Thunderbird, etc.) is selected.

### **2. Test the mailto handler directly**
Open **Run (Win+R)** and enter:
```
mailto:test@example.com
```
If nothing happens, the system’s `mailto:` association is broken.

### **3. Shorten or encode the email body**
Long bodies or unencoded characters can break the URL.  
Try:
- Reducing the body length  
- Ensuring proper URL encoding (spaces → `%20`, line breaks → `%0D%0A`)

### **4. Restart or repair the mail client**
Some clients (especially Outlook) may need:
- A restart  
- A profile repair  
- A system reboot  

### **5. Check corporate or system restrictions**
On managed devices, protocol handlers may be disabled.  
Contact IT if:
- `mailto:` does nothing  
- You receive **Access Denied** errors  

### **6. Try the built‑in webmail fallback**
If the desktop handler fails, the procedure automatically falls back to:
- Gmail  
- Outlook.com  

This ensures the user can still compose an email even when the local handler is misconfigured.

---
