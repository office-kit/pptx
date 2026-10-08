---
'@office-kit/pptx-editor': patch
'@office-kit/pptx-dev': patch
---

The editor and the dev tool no longer show third-party product names or our own branding in their UI.

- **Editor.** The title bar starts with the File menu; the "◈ @office-kit/pptx Editor" mark is gone, so an embedded editor carries no brand. The Help menu's product-named help item is now "Editor Help" (エディター ヘルプ), the Tools menu's product-named add-ins item is "Add-ins..." (アドイン), and Share ▸ Send a Copy offers "PPTX Presentation". Tooltips for unavailable features describe what is missing (for example "Translation needs an online translation service.", "Macros (VBA) do not run in this editor.", "Soft edges are not available for text.") instead of naming a third-party product or service.
- **Dev tool.** The preview, presenter and editor pages are titled "Presentation preview", "Presenter view" and "Presentation editor", and the scaffolded project instructions and agent prompts say "presentation" rather than naming a product.
