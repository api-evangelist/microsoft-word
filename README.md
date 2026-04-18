# Microsoft Word (microsoft-word)
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
