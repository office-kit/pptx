---
'@office-kit/pptx': patch
---

Fixed: `getCoreProperties` returned `null`, and `setCoreProperties` added a second core-properties part and relationship, for packages that store core properties somewhere other than `/docProps/core.xml` (e.g. `package/services/metadata/core-properties/….psmdcp`) (#457). Core and extended properties are now found through the package's root relationships, as OPC specifies, so `getExtendedProperties` / `setExtendedProperties` also work when `app.xml` lives elsewhere.
