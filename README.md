# Microsoft Word (microsoft-word)

<!-- API-EVANGELIST-PROVENANCE:BEGIN -->
> ### About this repository
>
> **This is not our API.** This repository is an independent, third-party profile of a company's
> **publicly available** API surface, maintained by [API Evangelist](https://apievangelist.com).
> API Evangelist does not operate, host, resell, or support this company's APIs, and is not
> affiliated with or endorsed by the company unless stated on the profile.
>
> **Where the information came from.** Everything here is assembled from material a member of the
> public can reach with a browser and no credentials — the company's own website, developer portal
> and documentation, the specifications it publishes for public use (OpenAPI, AsyncAPI, JSON Schema,
> `apis.json`, `llms.txt` and similar), its public repositories, and its public status, pricing and
> changelog pages. **Nothing here is obtained by breaching a system, defeating an access control, or
> using credentials of any kind.**
>
> **The rating is an independent assessment.** The Kin Score and Agent Readiness rating are
> independently calculated scores of a company's *public* API artifacts, produced by API Evangelist
> against a published rubric. They are not certifications, endorsements, security assessments, or
> audits, and they score published artifacts — not the quality, safety, or security of the software.
>
> **Corrections, re-scores, and removal are free.** No partnership, contract, or purchase is
> required, and you do not need to justify the request.
>
> - **Something wrong?** Open an issue on this repository, or email
>   [info@apievangelist.com](mailto:info@apievangelist.com).
> - **Published something new?** Ask for a re-score and we will re-run the rating.
> - **Want the listing taken down?** Say so and we will honor it. The profile is reduced to your
>   company name, a factual description, and a link to your own site, and the company is recorded as
>   **unrated** — never scored zero for having asked.
>
> **Response times.** Acknowledgement within **one business day**; removal or restriction within
> **two business days**; corrections and re-scores within **five business days**.
>
> **Not from the company, and here with a question?** You are welcome here — we would rather be the
> front line and point you the right way than have a good report go nowhere. What this repository
> can answer is narrow, though, so it is worth knowing who you are actually looking for:
>
> - **A question about how the API works, an account, billing, or a bug in the service** — that is
>   the company's own support, not us. We profile this API; we do not operate it and cannot see
>   your account.
> - **A bug in an open-source project we only catalog** — file it on that project's own repository.
>   This has happened with a real and correct bug report that reached us instead of the people who
>   could fix it, which helped nobody.
> - **Anything about this listing itself** — the description, the tags, the rating, a missing or
>   wrong artifact — is ours. Open an issue here.
> - **Not sure, or something general about API Evangelist or APIs.io** — open an issue on the
>   [APIs.io Inbox](https://github.com/api-search/inbox) and we will route it.
>
> **This repository contains no software, and we will never ask you to download anything.** There is
> no build, release, installer, or binary here — only text and machine-readable API descriptions, so
> there is nothing here that can be "corrupt" or need "repairing". Any issue, comment, or email
> claiming otherwise and offering a download link is not from us and is hostile. Do not follow the
> link; it is a lure. Report it to GitHub and, if you like, tell us at
> [info@apievangelist.com](mailto:info@apievangelist.com) so we can take it down.
>
> **On a security or compliance team?** Email
> [info@apievangelist.com](mailto:info@apievangelist.com) with *security* in the subject line and
> you will get a person, not a form. We will tell you exactly which public URLs this profile was
> built from so your team can see the same surface we did, and we will take the listing down on
> request while you work through it.
>
> Full detail: **[Where this data comes from](https://apievangelist.com/about/where-our-data-comes-from)**
<!-- API-EVANGELIST-PROVENANCE:END -->

APIs for Microsoft Word document creation, manipulation, and automation across Microsoft 365 cloud services, Office Add-ins, SharePoint, and Open XML document processing.

**URL:** [Visit APIs.json URL](https://raw.githubusercontent.com/api-evangelist/microsoft-word/refs/heads/main/apis.yml)

**Run:** [Capabilities Using Naftiko](https://github.com/naftiko/fleet?utm_source=api-evangelist&utm_medium=readme&utm_campaign=company-api-evangelist&utm_content=repo)

## Tags:

 - Documents, Microsoft 365, Office, Productivity, Word Processing

## Timestamps

- **Created:** 2024
- **Modified:** 2026-04-18

## APIs

### Microsoft Graph Word API
REST API for interacting with Word documents in Microsoft 365 and OneDrive via the Microsoft Graph unified endpoint. Provides operations for file management, content access, sharing, permissions, and document metadata.

**Human URL:** [https://learn.microsoft.com/en-us/graph/api/overview?view=graph-rest-1.0](https://learn.microsoft.com/en-us/graph/api/overview?view=graph-rest-1.0)

#### Tags:

 - Cloud, Documents, Microsoft Graph, REST

#### Properties

- [Documentation](https://learn.microsoft.com/en-us/graph/api/overview?view=graph-rest-1.0)
- [OpenAPI](openapi/microsoft-word-graph-api.yaml)
- [Authentication](https://learn.microsoft.com/en-us/graph/auth/)
- [APIReference](https://learn.microsoft.com/en-us/graph/api/resources/driveitem?view=graph-rest-1.0)
- [C# SDK](https://www.nuget.org/packages/Microsoft.Graph)
- [Python SDK](https://pypi.org/project/msgraph-sdk/)
- [JavaScript SDK](https://www.npmjs.com/package/@microsoft/microsoft-graph-client)
- [Java SDK](https://github.com/microsoftgraph/msgraph-sdk-java)
- [Go SDK](https://github.com/microsoftgraph/msgraph-sdk-go)
- [PHP SDK](https://github.com/microsoftgraph/msgraph-sdk-php)
- [DriveItem Schema](json-schema/graph-api-drive-item-schema.json)
- [Permission Schema](json-schema/graph-api-permission-schema.json)
- [DriveItem Structure](json-structure/graph-api-drive-item-structure.json)
- [Permission Structure](json-structure/graph-api-permission-structure.json)
- [JSON-LD Context](json-ld/microsoft-word-graph-api-context.jsonld)
- [DriveItem Example](examples/graph-api-drive-item-example.json)
- [Permission Example](examples/graph-api-permission-example.json)

### Office JavaScript API for Word
JavaScript API for building Word add-ins and automating Word tasks. Provides strongly-typed objects for document manipulation, content controls, tables, formatting, comments, collaboration, and shapes.

**Human URL:** [https://learn.microsoft.com/en-us/office/dev/add-ins/reference/overview/word-add-ins-reference-overview](https://learn.microsoft.com/en-us/office/dev/add-ins/reference/overview/word-add-ins-reference-overview)

#### Tags:

 - Add-Ins, Automation, Client-Side, JavaScript

#### Properties

- [Documentation](https://learn.microsoft.com/en-us/javascript/api/word)
- [OpenAPI](openapi/microsoft-word-javascript-api.yaml)
- [GettingStarted](https://learn.microsoft.com/en-us/office/dev/add-ins/quickstarts/word-quickstart)
- [CodeExamples](https://github.com/OfficeDev/Office-Add-in-samples)
- [APIReference](https://learn.microsoft.com/en-us/javascript/api/word?view=word-js-preview)
- [Tutorials](https://learn.microsoft.com/en-us/office/dev/add-ins/tutorials/word-tutorial)
- [Paragraph Schema](json-schema/javascript-api-paragraph-schema.json)
- [ContentControl Schema](json-schema/javascript-api-content-control-schema.json)
- [Table Schema](json-schema/javascript-api-table-schema.json)
- [Comment Schema](json-schema/javascript-api-comment-schema.json)
- [Paragraph Structure](json-structure/javascript-api-paragraph-structure.json)
- [ContentControl Structure](json-structure/javascript-api-content-control-structure.json)
- [Table Structure](json-structure/javascript-api-table-structure.json)
- [Comment Structure](json-structure/javascript-api-comment-structure.json)
- [JSON-LD Context](json-ld/microsoft-word-javascript-api-context.jsonld)
- [Paragraph Example](examples/javascript-api-paragraph-example.json)
- [ContentControl Example](examples/javascript-api-content-control-example.json)
- [Table Example](examples/javascript-api-table-example.json)
- [Comment Example](examples/javascript-api-comment-example.json)

### Word Automation Services (SharePoint)
Server-side document conversion and automation service for SharePoint. Supports batch conversion of Word documents to PDF, XPS, and other formats without user interaction.

**Human URL:** [https://learn.microsoft.com/en-us/sharepoint/dev/general-development/word-automation-services-in-sharepoint](https://learn.microsoft.com/en-us/sharepoint/dev/general-development/word-automation-services-in-sharepoint)

#### Tags:

 - Conversion, Enterprise, Server-Side, SharePoint

#### Properties

- [Documentation](https://learn.microsoft.com/en-us/sharepoint/dev/general-development/word-automation-services-in-sharepoint)

### Open XML SDK for Word
.NET library for programmatically creating and manipulating Word documents using the ECMA-376 Open XML standard. Provides strongly-typed classes for document structure, styles, tables, images, and content manipulation.

**Human URL:** [https://learn.microsoft.com/en-us/office/open-xml/open-xml-sdk](https://learn.microsoft.com/en-us/office/open-xml/open-xml-sdk)

#### Tags:

 - Dotnet, Library, Offline, OpenXML

#### Properties

- [Documentation](https://learn.microsoft.com/en-us/office/open-xml/open-xml-sdk)
- [OpenAPI](openapi/microsoft-word-open-xml-sdk.yaml)
- [GitHubRepository](https://github.com/OfficeDev/Open-XML-SDK)
- [.NET SDK](https://www.nuget.org/packages/DocumentFormat.OpenXml)
- [DocumentProperties Schema](json-schema/open-xml-sdk-document-properties-schema.json)
- [DocumentProperties Structure](json-structure/open-xml-sdk-document-properties-structure.json)
- [JSON-LD Context](json-ld/microsoft-word-open-xml-sdk-context.jsonld)
- [DocumentProperties Example](examples/open-xml-sdk-document-properties-example.json)

## Common Properties

- [Portal](https://developer.microsoft.com/)
- [DeveloperPortal](https://developer.microsoft.com/en-us/graph)
- [Console](https://developer.microsoft.com/en-us/graph/graph-explorer)
- [SignUp](https://developer.microsoft.com/en-us/microsoft-365/dev-program)
- [Authentication](https://learn.microsoft.com/en-us/graph/auth/)
- [GettingStarted](https://learn.microsoft.com/en-us/graph/get-started)
- [Blog](https://devblogs.microsoft.com/microsoft365dev/)
- [Support](https://developer.microsoft.com/graph/support)
- [StatusPage](https://status.office.com/)
- [ChangeLog](https://developer.microsoft.com/en-us/graph/changelog)
- [TermsOfService](https://www.microsoft.com/en-us/legal/terms-of-use)
- [PrivacyPolicy](https://privacy.microsoft.com/)
- [GitHubOrganization](https://github.com/OfficeDev)
- [StackOverflow](https://stackoverflow.com/questions/tagged/microsoft-graph)
- [X](https://twitter.com/MSGraphDev)
- [YouTube](https://www.youtube.com/@Microsoft365Developer)
- [Training](https://learn.microsoft.com/en-us/training/browse/?products=ms-graph)

## Features

| Name | Description |
|------|-------------|
| Document Creation | Programmatically create Word documents from templates or scratch using REST APIs or JavaScript. |
| Document Conversion | Convert Word documents to PDF, HTML, and other formats using server-side automation services. |
| Content Manipulation | Insert, edit, and format text, paragraphs, tables, images, and content controls in Word documents. |
| Collaboration | Track changes, manage comments, co-authoring sessions, and revision history through APIs. |
| Template Processing | Mail merge, document assembly, and template-based document generation for enterprise workflows. |
| Cloud Storage Integration | Access and manipulate Word documents stored in OneDrive, SharePoint, and Microsoft 365 cloud services. |
| Add-In Extensibility | Build custom Word add-ins with JavaScript APIs for task panes, content insertion, and document automation. |
| Open XML Processing | Low-level manipulation of Word document structure using the ECMA-376 Open XML standard. |

## Use Cases

| Name | Description |
|------|-------------|
| Automated Report Generation | Generate standardized reports from data sources using templates and API-driven document creation. |
| Contract and Legal Document Assembly | Assemble legal documents by merging clauses, terms, and client data into Word templates. |
| Document Review Workflows | Automate review cycles with tracked changes, comments, and approval workflows via APIs. |
| Bulk Document Processing | Process large volumes of Word documents for format conversion, content extraction, or metadata updates. |
| Custom Business Add-Ins | Build task pane add-ins that connect Word to CRM, ERP, or other business systems for data insertion. |
| Compliance Document Management | Ensure regulatory compliance by automating document formatting, metadata tagging, and archiving. |

## Integrations

| Name | Description |
|------|-------------|
| Microsoft SharePoint | Store, manage, and collaborate on Word documents through SharePoint document libraries and workflows. |
| Microsoft OneDrive | Access and sync Word documents via OneDrive cloud storage through Microsoft Graph APIs. |
| Microsoft Teams | Collaborate on Word documents directly within Teams channels and chat conversations. |
| Microsoft Power Automate | Automate Word document workflows including creation, conversion, and approval routing with Power Automate flows. |
| Microsoft Copilot | AI-powered document creation, editing, and summarization through Copilot agents with add-in actions. |
| Azure Active Directory | OAuth 2.0 authentication and authorization for secure API access through Microsoft Identity Platform. |

## Artifacts

Machine-readable API specifications organized by format.

### OpenAPI

- [Microsoft Graph Word API](openapi/microsoft-word-graph-api.yaml)
- [Microsoft Word JavaScript API](openapi/microsoft-word-javascript-api.yaml)
- [Microsoft Word Open XML SDK](openapi/microsoft-word-open-xml-sdk.yaml)

### JSON Schema

- [DriveItem Schema](json-schema/graph-api-drive-item-schema.json)
- [Permission Schema](json-schema/graph-api-permission-schema.json)
- [Paragraph Schema](json-schema/javascript-api-paragraph-schema.json)
- [ContentControl Schema](json-schema/javascript-api-content-control-schema.json)
- [Table Schema](json-schema/javascript-api-table-schema.json)
- [Comment Schema](json-schema/javascript-api-comment-schema.json)
- [DocumentProperties Schema](json-schema/open-xml-sdk-document-properties-schema.json)

### JSON Structure

- [DriveItem Structure](json-structure/graph-api-drive-item-structure.json)
- [Permission Structure](json-structure/graph-api-permission-structure.json)
- [Paragraph Structure](json-structure/javascript-api-paragraph-structure.json)
- [ContentControl Structure](json-structure/javascript-api-content-control-structure.json)
- [Table Structure](json-structure/javascript-api-table-structure.json)
- [Comment Structure](json-structure/javascript-api-comment-structure.json)
- [DocumentProperties Structure](json-structure/open-xml-sdk-document-properties-structure.json)

### JSON-LD

- [Graph API Context](json-ld/microsoft-word-graph-api-context.jsonld)
- [JavaScript API Context](json-ld/microsoft-word-javascript-api-context.jsonld)
- [Open XML SDK Context](json-ld/microsoft-word-open-xml-sdk-context.jsonld)

### Examples

- [DriveItem Example](examples/graph-api-drive-item-example.json)
- [Permission Example](examples/graph-api-permission-example.json)
- [Paragraph Example](examples/javascript-api-paragraph-example.json)
- [ContentControl Example](examples/javascript-api-content-control-example.json)
- [Table Example](examples/javascript-api-table-example.json)
- [Comment Example](examples/javascript-api-comment-example.json)
- [DocumentProperties Example](examples/open-xml-sdk-document-properties-example.json)

## Capabilities

Naftiko capabilities organized as shared per-API definitions composed into customer-facing workflows.

### Shared Per-API Definitions

- [Microsoft Graph Word API](capabilities/shared/graph-api.yaml) -- 8 operations for cloud document management
- [Office JavaScript API for Word](capabilities/shared/javascript-api.yaml) -- 9 operations for document content manipulation
- [Open XML SDK for Word](capabilities/shared/open-xml-sdk.yaml) -- 6 operations for server-side document processing

### Workflow Capabilities

| Workflow | APIs Combined | Tools | Persona |
|----------|--------------|-------|---------|
| [Document Management](capabilities/document-management.yaml) | Graph API + JavaScript API + Open XML SDK | 21 | Document Author, Content Manager, Automation Engineer |

## Vocabulary

- [Microsoft Word Vocabulary](vocabulary/microsoft-word-vocabulary.yaml) -- Unified taxonomy mapping 22 resources, 14 actions, 1 workflow, and 3 personas across operational (OpenAPI) and capability (Naftiko) dimensions

## Rules

- [Microsoft Word Spectral Rules](rules/microsoft-word-spectral-rules.yml) -- 35 rules across 12 categories enforcing Microsoft Word API conventions

## Maintainers

**FN:** Kin Lane

**Email:** kin@apievangelist.com
