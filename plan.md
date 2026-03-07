# SharePoint Migration Planner — Implementation Plan

## Overview
A **browser-based dashboard** (static SPA) that helps plan and orchestrate SharePoint Online migrations. The user signs in with their M365 account, uploads a TreeSize report, and visually maps file system nodes to SharePoint sites and document libraries. If target sites don't exist, the dashboard can create them.

**Architecture**: Static SPA (TypeScript + Vite) → Azure Static Web Apps (free) → Microsoft Graph API
**Auth**: Delegated (MSAL.js + PKCE) — user signs in via browser, no secrets in code
**Hosting**: Azure Static Web Apps free tier, deployed in the client's Azure tenant

---

## Phase 1: Project Scaffolding

### 1.1 Initialize Vite + TypeScript project
```bash
npm create vite@latest sp-migration-planner -- --template vanilla-ts
```
- Vite for dev server + production build (outputs static files)
- TypeScript for type safety
- No framework initially — vanilla TS with DOM manipulation (lightweight)
- Can add a lightweight UI library later if needed (e.g., Lit, Preact)

### 1.2 Directory structure
```
sp-migration-planner/
├── src/
│   ├── main.ts                    # App entry point
│   ├── auth/
│   │   ├── msalConfig.ts          # MSAL browser configuration
│   │   └── authService.ts         # Sign-in/out, token acquisition
│   ├── graph/
│   │   └── graphClient.ts         # Graph API wrapper (sites, drives)
│   ├── parsers/
│   │   ├── treeSizeParser.ts      # TreeSize CSV/Excel parser
│   │   └── fileDetector.ts        # Auto-detect CSV vs Excel
│   ├── ui/
│   │   ├── components/
│   │   │   ├── authPanel.ts       # Sign-in/sign-out UI
│   │   │   ├── uploadPanel.ts     # File upload drop zone
│   │   │   ├── treeView.ts        # TreeSize data as interactive tree
│   │   │   ├── mappingPanel.ts    # Drag/drop or select mapping UI
│   │   │   ├── siteCreator.ts     # Create new sites form
│   │   │   └── summaryPanel.ts    # Migration plan summary/export
│   │   └── app.ts                 # Main app shell / layout
│   ├── state/
│   │   └── store.ts               # Simple state management
│   └── types/
│       └── index.ts               # Shared TypeScript interfaces
├── index.html                     # Single HTML entry point
├── public/
│   └── favicon.svg
├── package.json
├── tsconfig.json
├── vite.config.ts
└── staticwebapp.config.json       # Azure Static Web Apps config
```

### 1.3 Dependencies
| Package | Purpose |
|---|---|
| `@azure/msal-browser` | Delegated auth (PKCE flow) in the browser |
| `@microsoft/microsoft-graph-client` | Graph API calls for SharePoint |
| `exceljs` | Parse TreeSize Excel exports (works in browser) |
| `papaparse` | Parse TreeSize CSV exports (browser-native, lightweight) |

Dev dependencies: `vite`, `typescript`

> **Note**: No backend dependencies. No Express, no dotenv, no winston. Everything runs client-side.

---

## Phase 2: Authentication (MSAL.js Delegated)

### 2.1 App Registration Setup (manual step, documented in README)
In the client's Azure portal:
1. Register a new app (e.g., "SP Migration Planner")
2. Set **SPA** redirect URI: `http://localhost:5173` (dev) + production URL
3. Add **delegated** permissions:
   - `Sites.ReadWrite.All` — browse and manage SharePoint sites
   - `Sites.Manage.All` — create new site collections (if needed)
   - `User.Read` — basic profile for the signed-in user
4. **No client secret needed** — SPA uses PKCE

### 2.2 MSAL Configuration (`src/auth/msalConfig.ts`)
```typescript
// Configuration object — client ID and tenant ID are the only required values
{
  auth: {
    clientId: "<from-env-or-config>",
    authority: "https://login.microsoftonline.com/<tenant-id>",
    redirectUri: window.location.origin
  }
}
```

### 2.3 Auth Service (`src/auth/authService.ts`)
- `signIn()` — popup or redirect login
- `signOut()` — clear session
- `getToken(scopes)` — acquire token silently (with fallback to interactive)
- `isAuthenticated()` — check current auth state
- Expose the signed-in user's display name and tenant for the UI header

---

## Phase 3: TreeSize Report Parsing

### 3.1 TreeSize Export Format
TreeSize Pro/Free exports typically contain:
| Column | Description |
|---|---|
| `Path` | Full folder/file path (e.g., `\\server\share\Projects\2024`) |
| `Size` | Size in bytes or human-readable |
| `Files` | Number of files in folder |
| `Folders` | Number of subfolders |
| `% of Parent` | Percentage of parent folder size |
| `Last Change` | Last modified date |

### 3.2 Parser Implementation (`src/parsers/treeSizeParser.ts`)
- Accept uploaded `File` object from browser file input
- Auto-detect CSV (`.csv`) vs Excel (`.xlsx`) via extension
- **CSV**: Use `papaparse` — browser-native, streaming capable
- **Excel**: Use `exceljs` — load workbook from `ArrayBuffer`
- Normalize data into a typed `TreeNode[]` structure:
  ```typescript
  interface TreeNode {
    path: string;           // Original file system path
    name: string;           // Folder/file name
    depth: number;          // Nesting level
    sizeBytes: number;      // Size in bytes
    fileCount: number;      // Number of files
    folderCount: number;    // Number of subfolders
    lastModified?: Date;    // Last change date
    children: TreeNode[];   // Nested children (built from paths)
  }
  ```
- Build a **tree structure** from flat path data (split paths, nest by hierarchy)

### 3.3 Validation
- Check for required columns (`Path` at minimum)
- Handle different TreeSize versions/export formats gracefully
- Show warnings for unrecognized columns

---

## Phase 4: Dashboard UI

### 4.1 App Layout (`src/ui/app.ts`)
A single-page layout with a step-based workflow:

```
┌─────────────────────────────────────────────────────┐
│  SP Migration Planner          [User Name] [Logout] │
├─────────────────────────────────────────────────────┤
│  Step 1: Connect  │  Step 2: Upload  │  Step 3: Map │
├────────────────────────┬────────────────────────────┤
│                        │                            │
│   [Active Panel]       │   [Context Panel]          │
│                        │                            │
└────────────────────────┴────────────────────────────┘
```

### 4.2 Step 1 — Connect (Auth Panel)
- "Sign in with Microsoft" button
- Shows tenant name + user after sign-in
- Tests Graph API connectivity (fetches root site)
- Status indicator: connected / disconnected

### 4.3 Step 2 — Upload TreeSize Report (Upload Panel)
- Drag-and-drop zone or file picker
- Accepts `.csv` and `.xlsx` files
- Parses on upload and shows preview:
  - Total folders/files detected
  - Total size
  - Root path
- Renders parsed data as an **interactive tree view** (expandable/collapsible)

### 4.4 Step 3 — Map to SharePoint (Mapping Panel)
This is the core of the dashboard. Split view:

```
┌─ TreeSize Structure ──────────┬─ SharePoint Target ────────────┐
│                               │                                │
│  ▼ \\Server\Share             │                                │
│    ▼ Projects                 │  → contoso.sharepoint.com/     │
│      ▼ Engineering            │      sites/Engineering         │
│        📁 Docs (45 GB)  [Map]│      └─ Documents/Docs         │
│        📁 Data (12 GB)  [Map]│      └─ Documents/Data         │
│      ▼ Marketing              │    sites/Marketing             │
│        📁 Assets (8 GB) [Map]│      └─ Documents/Assets       │
│    ▼ Archive                  │  → (not mapped)                │
│                               │                                │
└───────────────────────────────┴────────────────────────────────┘
```

**Mapping workflow**:
1. User selects a TreeSize node (folder)
2. Right panel shows:
   - **Existing sites**: Fetched from Graph API (`GET /sites?search=*`), filterable
   - **Existing document libraries**: Fetched for the selected site
   - **"Create new site" option**: Opens site creation form
3. User picks a target site + document library + optional subfolder path
4. Mapping is saved in state and shown visually

### 4.5 Step 3b — Create New Sites (Site Creator)
When a target site doesn't exist:
- Form fields: Site name, URL suffix, description, template (Team site / Communication site)
- Calls Graph API: `POST /sites` (or SharePoint admin API)
- Shows creation status (pending / created / failed)
- Newly created site becomes available in the mapping dropdown

### 4.6 Step 4 — Summary & Export (Summary Panel)
- Table view of all mappings:
  | Source Path | Size | Files | Target Site | Target Library | Target Path | Status |
  |---|---|---|---|---|---|---|
  | `\\Server\Projects\Eng\Docs` | 45 GB | 1,200 | Engineering | Documents | /Docs | Ready |
- Totals: total data size, total files, number of sites to create
- **Export mapping plan** as CSV or JSON (for use by a future migration execution tool)
- Highlights unmapped nodes as warnings

---

## Phase 5: Graph API Integration

### 5.1 Graph Client (`src/graph/graphClient.ts`)
Wrapper methods using `@microsoft/microsoft-graph-client`:

| Method | Graph API Call | Purpose |
|---|---|---|
| `getCurrentUser()` | `GET /me` | Display signed-in user |
| `getRootSite()` | `GET /sites/root` | Verify connectivity + get tenant |
| `searchSites(query)` | `GET /sites?search={query}` | Search for existing sites |
| `getSiteDrives(siteId)` | `GET /sites/{id}/drives` | List document libraries |
| `getDriveContents(driveId, path)` | `GET /drives/{id}/root:/{path}:/children` | Browse library contents |
| `createTeamSite(name, alias)` | `POST /groups` (M365 group-connected) | Create team site |
| `createCommSite(name, url)` | SharePoint REST API or Graph beta | Create communication site |

### 5.2 Site Creation Details
- **Team sites**: Created via M365 Group (`POST /groups` with `groupTypes` including `Unified`)
- **Communication sites**: May require SharePoint-specific API (`/_api/SPSiteManager/create`) called via Graph proxy or direct REST
- Poll for provisioning completion before showing as "ready"

---

## Phase 6: State Management

### 6.1 Simple Store (`src/state/store.ts`)
No framework needed — a simple observable store pattern:

```typescript
interface AppState {
  auth: { user: User | null; isAuthenticated: boolean };
  treeData: TreeNode | null;           // Parsed TreeSize data
  mappings: MigrationMapping[];        // Source → destination mappings
  sites: SharePointSite[];             // Cached site list from Graph
  pendingSiteCreations: SiteRequest[]; // Sites queued for creation
}
```

- State changes trigger UI re-renders for affected panels
- State is persisted to `sessionStorage` so a browser refresh doesn't lose work

---

## Phase 7: Hosting & Deployment

### 7.1 Azure Static Web Apps
- `staticwebapp.config.json` for SPA routing (fallback to `index.html`)
- Deploy via GitHub Actions (auto-configured by Azure) or Azure CLI
- Free tier includes: custom domain, SSL, global CDN, auth integration
- No backend, no Azure Functions needed for Phase 1

### 7.2 Local Development
- `npm run dev` — Vite dev server on `localhost:5173`
- Hot module reload for fast iteration
- App Registration redirect URI includes `http://localhost:5173` for local auth

---

## Build Order (Implementation Sequence)

| Step | What | Depends on | Deliverable |
|---|---|---|---|
| 1 | Vite + TS scaffolding, dependencies | — | Empty app builds and runs |
| 2 | Types / interfaces | — | `types/index.ts` |
| 3 | MSAL config + auth service | Step 1 | Sign-in/out works |
| 4 | Auth panel UI (sign-in button, user display) | Step 3 | User can authenticate |
| 5 | Graph client wrapper | Step 3 | Can call Graph API |
| 6 | TreeSize CSV parser | Step 2 | Parses CSV to TreeNode[] |
| 7 | TreeSize Excel parser | Step 2 | Parses Excel to TreeNode[] |
| 8 | Upload panel + tree view UI | Step 6, 7 | User can upload and see tree |
| 9 | Site search + library listing (Graph) | Step 5 | Can browse existing SharePoint sites |
| 10 | Mapping panel UI | Step 8, 9 | User can map nodes to sites |
| 11 | Site creation (Graph) | Step 5 | Can create new sites |
| 12 | Site creator UI | Step 11 | User can create sites from dashboard |
| 13 | Summary panel + export | Step 10 | User can review and export plan |
| 14 | State persistence (sessionStorage) | Step 10 | Survives page refresh |
| 15 | Azure Static Web Apps config + deploy | Step 13 | Deployed to Azure |

---

## Out of Scope (Future Phases)
- **Migration execution** — actually moving files (Phase 2 of the larger project)
- **Progress tracking** — real-time migration progress dashboard
- **Permission mapping** — mapping NTFS permissions to SharePoint permissions
- **Metadata mapping** — preserving file metadata during migration
- **Incremental sync** — detecting changes and migrating only deltas
- **Multi-tenant support** — currently single-tenant per deployment
- **Azure Function backend** — only needed if app-only auth is required later
