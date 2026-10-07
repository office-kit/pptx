---
'@office-kit/pptx-editor': patch
'@office-kit/pptx-dev': patch
---

The editor and the dev tool no longer show third-party product names or our own branding in their UI.

- **Editor.** The title bar starts with the File menu; the "◈ @office-kit/pptx Editor" mark is gone, so an embedded editor carries no brand. Help ▸ "PowerPoint Help" is now "Editor Help" (エディター ヘルプ), Tools ▸ "PowerPoint Add-ins..." is "Add-ins..." (アドイン), and Share ▸ "Send a Copy (PowerPoint Presentation)" is "Send a Copy (PPTX Presentation)". Tooltips for unavailable features describe what is missing (for example "Translation needs an online translation service.", "Macros (VBA) do not run in this editor.", "Soft edges are not available for text.") instead of naming Microsoft, Microsoft 365, Office, OneDrive, SharePoint or PowerPoint.
- **Dev tool.** The preview, presenter and editor pages are titled "Presentation preview", "Presenter view" and "Presentation editor", and the scaffolded project instructions and agent prompts say "presentation" rather than naming a product.
