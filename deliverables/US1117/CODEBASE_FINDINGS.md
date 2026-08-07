# Guard onboarding — what the code actually says

Read of three repos: **SAM-OnSite** (Flutter app), **SAM** (Python AI backend), **Pronect-Management-API** (.NET). Every claim below has a file path. Where the code does not settle a question, it says so.

---

## 1. How the system fits together

| Piece | What it is | Role in onboarding |
|---|---|---|
| **SAM-OnSite** | Flutter app, Supabase for auth and data | Where onboarding is built. Also the **only writer** of guard activities. |
| **SAM** | FastAPI AI backend | **Reads** operational data over MCP and answers questions. Creates no reports. |
| **OnSite MCP server** | Not in any repo | Where activities are actually queried. The exclusion filter's real home. |
| **Pronect-Management-API** | Document ingestion (`MetadataType`, `MetadataContent`) | Not involved. |

Activities live in a Supabase `events` table. One write path: `SAMOnSite/lib/data/repositories/event_repository.dart:75` `upsert(Event)`.

When the product owner said *"update the reporting layer"*, that layer is the **OnSite MCP server** plus SAM's rendering — not the Management API.

---

## 2. Five things the specification assumed that the code contradicts

### 2.1 Press-and-hold to edit a message does not exist

Zero `onLongPress` in `lib/`. `GestureDetector` appears twice, both `onTap` on photos. `_MessageBubble` (`conversation_page.dart:2252`) has no gesture handler.

The data layer supports it — `message_repository.dart:108 upsertMessage`, `Message.previousMessageId` — but no UI calls it.

**This contradicts a tested answer of "Yes"** and the train-the-trainer deck's tip 6. Either the tested build differs from this snapshot, the feature is in the Pronect app, or something else was tested. **Resolve before build**: it is one of the two correction steps onboarding teaches.

### 2.2 NFC does not exist in SAM OnSite

Zero matches for `nfc` across all three repos. No package, no manifest entry, no code.

So SAM OnSite cannot read NFC state, cannot detect NFC hardware, and never uses NFC. The gate can still send a guard to the setting, but the app has no way to know it worked, and nothing in this app depends on it. The checkpoint scanning is presumably the separate Pronect app.

### 2.3 The app already reads "allow all the time" — and throws it away

`SAMOnSite/lib/services/location_tracking_service.dart:30-42`:

```dart
final status  = await Permission.locationAlways.status;
final precise = await Permission.location.status;
final activity = await Permission.activityRecognition.status;
return locationGranted && preciseGranted && activityGranted;
```

Three states are read, then AND-ed into one boolean. `LocationTrackingState.permissionsGranted` (`location_controller.dart:27`) is the only thing the UI sees, so a guard who granted everything except always-on gets the same *"Permissions needed"* as one who granted nothing.

**The capability exists.** Surfacing which one is missing is small.

Also: only `.isGranted` is ever used. Nothing calls `isPermanentlyDenied`, so "denied once" and "never ask again" look identical — which matters for routing a guard to support.

### 2.4 Background tracking resets to off on every launch

`location_controller.dart:59-64` — `LocationTrackingState.initial()` hardcodes `trackingEnabled: false`. Nothing persists it. Nothing restarts it.

Onboarding turns it on, the guard restarts the phone, it is off, and nobody knows. **This is a live product defect, not an onboarding one.**

### 2.5 The archive flag is not in the app's Event model

`SAMOnSite/lib/data/models/event.dart` has thirteen fields and no archive, soft-delete or training flag.

**Careful:** the `events` table DDL is in no repo — there are **no Supabase migrations checked in anywhere** — so this does not prove the column is absent from the database. The reporting console shows soft-delete on activities, so it probably exists server-side.

What is certain: **the app cannot set it today.**

---

## 3. What already exists and can be used

| Capability | Where |
|---|---|
| A blocking gate, already exhaustive over phases | `SAMOnSite/lib/main.dart:98` `AuthGate` switching on `AppGatePhase` |
| The pattern for "don't reach ready until this succeeds" | `providers/auth_providers.dart` `_bootstrapTemplatesForUser` |
| Local key-value storage, already a dependency | `shared_preferences`, used in `services/template_service.dart` |
| Copy in English **and Dutch** | `localization/app_localizations.dart` — hand-rolled, one file |
| Permission reads including the always-on level | `services/location_tracking_service.dart:30-42` |
| The green mic (tip 1 is accurate) | `conversation_page.dart:1637` `MicPulseAnimation(color: Colors.green)` when `micOpen && !isSuspended` |
| Single event write path that auto-carries new fields | `event_repository.dart:75` — `toJson()` strips nulls, so a new field flows through |
| A working archive pattern to copy | `conversation_repository.dart:132` — `active: false` + `ended_at` |
| Consolidated read layer: 4 repo reads, 6 providers, 2 pages | `event_repository.dart`, `providers/chat_providers.dart`, `views/home/` |

There is **no onboarding, tutorial, coach-mark or first-run code anywhere**, not even dead code. This is greenfield.

---

## 4. How to build it

### Where onboarding goes

Add an `onboarding` phase to `AppGatePhase` between `loading` and `ready`. `AuthGate` is a `switch`, so the compiler flags what needs updating. It is blocking by construction and there is no router to fight — `MaterialApp` has only `home:`.

### Three risks in that placement

1. **The gate rebuilds on every resume.** `handleAppResumed()` (`auth_providers.dart:281`) refreshes the session; a network blip flips the phase to `reconnecting` or `degraded`. Combined with the permission-grant restart, a guard can be thrown back to step 1. **Persist progress per step, not at the end.**
2. **Boot is already fragile.** Reaching `ready` needs a session refresh *and* a template download with a 30s timeout that hard-fails offline. Decide explicitly whether onboarding may run offline.
3. **Permissions are never re-read on resume.** `refreshPermissions()` is wired only to a manual button, so a guard who grants everything and returns still sees *"Permissions needed"*.

### Sizing

**Small — one field or one line each**
- `is_training` on the `Event` model; it flows through `upsert` automatically
- Filter the four app-side reads: `watchByUser`, `watchByConversation`, `watchByConversationIds`, `getById`
- Onboarding stage in `SharedPreferences`
- Onboarding copy in `app_localizations.dart`
- Surface which permission is missing instead of one boolean

**Medium**
- Persist background tracking and restore it on launch
- Re-read permissions on resume
- Deep link to a specific settings page — today only `openAppSettings()` exists, which opens the app-info page. Anything more needs a package the app does not have (`android_intent_plus` or similar)
- A practice conversation. Every `ConversationPage` open inserts a real Supabase row (`conversation_page.dart:302`). There is no draft or sandbox mode
- Press-and-hold edit UI, if 2.1 confirms it is missing

**Large or unknown**
- **The exclusion filter in the OnSite MCP server.** Four tools read event data (`list_onsite_events`, `get_onsite_events_report_data`, `get_onsite_client_overview_report_data`, `get_onsite_control_round_report_data`). That server is in no repo. **This is the item gating go-live and it cannot be sized from what we have.**
- **A home for Supabase migrations.** None exist in version control. A new column has nowhere to live today.

---

## 5. One thing worth fixing first, cheaply

SAM does not allowlist MCP tools. It adopts whatever the server advertises:

```python
tool_spec = McpToolSpec(client=mcp_client)
tools = await tool_spec.to_tool_list_async()
```
`SAM_2/backend/pronect/pronect/nlp/onsite/onsite_mcp_tool_builder.py:436`

`KNOWN_ONSITE_TOOL_NAMES` exists but is only used for bookkeeping, not filtering. So the day the MCP server gains a write tool, the agent can call it with no change here.

An allowlist is cheap and it is the only structural guarantee available from this side that a practice conversation cannot reach a live write.

---

## 6. Questions this read answers

| Question | Answer |
|---|---|
| What is the physical activity permission for? | It gates location tracking (`location_tracking_service.dart:34,39,49`). **No code anywhere consumes activity data.** Denying it blocks tracking for a signal that is never read. Worth asking whether it is needed. |
| Can the app read whether NFC is on? | No. NFC does not exist in this app. |
| Can the app read permission state? | Yes, including the always-on location level. It reads it and discards it. |
| Does SAM need the notifications permission? | Declared (`AndroidManifest.xml:9`), never requested or read in Dart. The only consumer is geolocator's foreground-service notification. |
| Where is the practice flag set at creation? | Nowhere yet, but `event_repository.dart:75` is the single write path and would carry it in one operation. |

## 7. Questions this read cannot answer

- Whether the `events` table already has an archive column (no migrations in version control).
- What the OnSite MCP server does, or how many queries the exclusion filter must cover.
- Whether `RECORD_AUDIO` reaches the merged manifest via the `record` plugin.
- Whether the batched foreground+background location request actually yields "always" on a given Android version.
- Whether press-and-hold editing exists in a branch or app other than this snapshot.

---

# ADDENDUM — the Pronect platform, and why archive mode is not enough

Read of three further repos: **Pronect** (the platform, ~675 C# files), **sam-dab-mcp**, **Pronect-Function-App**.

## The finding that changes the plan

**Archiving an event does not stop the supervisor alert.**

`Pronect/PSM.PlatformEvent.DomainService/Services/PlatformEventService.cs:86-136`. The gate on the alert is whether the event is *new*, not whether it is active:

```csharp
var isNewPlatformEvent = !await platformEventRepository.ExistsAsync(...);
var platformEventIngested = await platformEventRepository.IngestAsync(platformEvent, userId);

if (isNewPlatformEvent)
{
    var recipientUsers = await GetEventEmailRecipientsAsync(...);
    _ = Task.Run(async () => {
        foreach (var recipientUser in recipientUsers)
            await emailService.SendCriticalPlatformEventEmailAsync(
                recipientUser.Email, recipientUser.FirstName,
                platformEventIngested.Title, platformEventIngested.Description, ...);
```

A training incident ingested with `Active = false` **still emails every subscribed supervisor** the made-up title and description. Fire-and-forget, outside the request scope, via SendGrid.

The whole safety design rests on "put it in archive mode". On this path, archive mode does nothing.

**The fix is one line** — add `&& platformEventIngested.Active` to the condition. It is the highest-value change in the entire piece of work.

## Where guard incidents actually live

`Pronect/PSM.Core/Entities/PlatformEvent.cs`. Ingested at `POST api/v1/PlatformEvent/IngestPlatformEvent` (`PSM.API/Controllers/PlatformEventController.cs:86`), a machine-to-machine endpoint with its own `PlatformEventIngestor` role. `SourceSystem` is documented as `SAM_ONSITE`; `SourceUserId` is a Supabase UUID, explicitly not a Pronect user.

**Careful:** `PSM.Core/Entities/Event.cs` is a different thing entirely — a risk-taxonomy catalogue row ("Theft", "Arson"). No relationship to `PlatformEvent`. Confusing the two would scope this work badly wrong.

## The archive flag exists

`PlatformEvent.Active`, bool, default true. Migration `20260727000000_AddPlatformEventActive.cs`. Set three ways: from the ingest payload at creation (so SAM OnSite can already create an archived event with no schema change), on upsert, and by `PATCH .../{id}/Archive`.

## The exclusion filter is far smaller than assumed

Only **one file** in the whole platform reads `PlatformEvent`: `PSM.Infrastructure/Data/Repositories/PlatformEventRepository.cs`. These services contain **zero references** to it:

- `PSM.Kpi.DomainService` — KPIs are scored against controls, not guard activity
- `PSM.Report.DomainService` — dashboards, RPC and the Power BI service read risk and control data
- `PSM.Risk.DomainService`
- `PSM.Location.DomainService`
- `PSM.Notification.DomainService` — resolves recipients, never sees the event

So four of the surfaces the specification lists as needing a filter **do not read this data at all**.

## Where the filter is genuinely needed, in order

| # | Place | Difficulty |
|---|---|---|
| 1 | `PlatformEventService.cs:86` — the supervisor email fires regardless of `Active` | **One line.** Do this first. |
| 2 | The SAM weekly/daily executive summary emailed to customers (`PSM.Email.DomainService/Services/SamEmailCampaignService.cs:878`) | **Cross-team, cannot be fixed here.** Content is generated by an external SAM API over its own data. Needs a contract change and work in SAM OnSite. Long lead time. |
| 3 | `PlatformEventRepository.cs:129 GetByIdAsync` — no `Active` filter, so a direct link renders an archived incident in full, including photo SAS URLs | One line |
| 4 | `ActionPlanServices.cs:158` — an action plan can be created against a training incident, which emails the owner and persists in the Action Tracker | Easy |
| 5 | `PlatformEventRepository.cs:222` — re-ingest overwrites `Active`, silently un-archiving | Design decision, see below |
| 6 | `PlatformEventRepository.cs:384 GetDistinctUsersAsync` — the training guard leaks into the reporter typeahead | One line |
| 7 | `PlatformEventParameters.cs:121 IncludeArchived` — a client-controlled parameter that unhides everything | See below |
| 8 | Audit log (`PlatformEventRepository.cs:203`, surfaced by `LogController`) keeps `Title` and `Description` | Medium |
| 9 | A global query filter in `PSMDbContext.cs:763` | Medium, and riskier than it looks |

## Do not reuse `Active` for training

Two reasons, both concrete:

- **Re-ingest resurrects it.** `PlatformEventRepository.cs:222` does `existing.Active = entity.Active;`. Anything re-sent with `Active=true` un-archives.
- **Clients can unhide it.** `IncludeArchived` is a query parameter with a default of false. Any caller can ask for archived rows. Fine for a user archiving their own event; wrong for training data that must never be visible.

**Recommendation: a separate `IsTraining` column, set at ingest, never updatable.** One migration, following the pattern already set by `AddPlatformEventActive`.

## If a global query filter is used, three things break

`PSMDbContext.cs` has no `HasQueryFilter` anywhere today. Adding one to `PlatformEvent` would break:

- `HasArchivedAsync` — its job is finding inactive rows
- The `IncludeArchived=true` path
- **`IngestAsync` and `ExistsAsync`** — the idempotency lookup would stop seeing archived events, so a re-ingest attempts an INSERT and violates the unique index on `(SourceClientId, SourceSystem, SourceEventId)`. That surfaces as a production 500, not a silent bug.

Each needs `.IgnoreQueryFilters()`. There is also no index on `Active` — add one for whichever flag is used.

## Two things that are not in any repo

- **The location heatmap.** Zero matches for `heatmap` across the platform. It is front-end or lives in SAM.
- **The scheduled executive summary generator.** Runs over the OnSite MCP path.

## The other two repos

**sam-dab-mcp** — a stock Data API Builder container exposing five read-only views (risk register, performance summary, KPI detail, cost performance, action plans), each scoped by `@item.client_id eq @claims.clientId`. Nothing writable in any config. None of them is guard activity, so no filter is needed unless practice data propagates upward into those aggregates.

**Pronect-Function-App** — eight functions, all timer or queue triggered, every outbound call a bodyless POST. It never reads data, so it cannot filter anything. If practice runs must not trigger the weekly SAM email or KPI notifications, that gate belongs in the services behind `SamEmailCampaignApi` and `KPIApi`.

## Still not found

Long-press-to-edit and NFC appear in **none of the six repos**, searched case-insensitively across every file type. The SAM-OnSite snapshot is `main` at version `1.3.8+2`. Since both features work on a real handset, the running build is not this snapshot — most likely a branch. An Azure DevOps code search for `onLongPress` will name it.

---

# ADDENDUM 2 — the live surfaces, finally enumerated

Read of **Pronect-Frontend** (Angular). This closes the question that has been open since the product owner's July message: which live surfaces must the exclusion filter cover.

Only two components in the whole app touch `PlatformEventService`. Everything else surfaces event data indirectly, which is where the leaks are.

## The twelve places a training incident could appear

Ordered by likelihood. "Inherits" means a backend `Active` filter on the existing query fixes it for free.

| # | Surface | Inherits? |
|---|---|---|
| 1 | **Reporter typeahead** — a practice-only guard's name and email appear in the Reporter suggestions for every user on the page | **No** |
| 2 | **Events table with the archived toggle on** | **Partially** |
| 3 | **Audit log** — event Title and Description in log rows and change summaries | **No** |
| 4 | **SAM chat location map** — pins carrying event title, type and description | **No** |
| 5 | **Action plan created from an event** — description copied into a persisted field | **No** |
| 6 | Detail modal on an archived row | No, but contained |
| 7 | Photos and their SAS URLs | Inherits |
| 8 | Linked action plans table | Inherits |
| 9 | SAM chat PDF export | Downstream of 4 |
| 10 | SAM chat stat tiles | Downstream of 4 |
| 11 | Pagination total count | Inherits |
| 12 | Power BI embedded pages | **Unknown** |

## The five that need their own change

### 1. The reporter typeahead is the leakiest surface

`src/app/services/platform-event.service.ts:69` calls `GET /v1/PlatformEvent/GetPlatformEventUsers`. The params interface has **no field for `includeArchived`**:

```ts
export interface IPlatformEventUserLookupParams {
  searchTerm: string;
  pageSize?: number;
  minStartedAt?: string;
  maxStartedAt?: string;
  locationId?: number;
  sourceSystem?: string;
}
```

The response merges emails, users **and full names**. So a guard who has only ever filed practice incidents still appears in the Reporter dropdown for anyone on the page, and the archived toggle makes no difference. Fixing the list query does not fix this.

### 2. The archived toggle is not role-gated

`platform-events-list.component.html:145` — the switch appears whenever `hasArchivedEvents` is true, which is data-driven, not permission-driven. Every role that can open the page gets it, including `ReadOnlyUser`. Archived rows are dimmed to 0.6 opacity and carry an archive icon, but remain fully clickable and open the complete detail modal.

This is the second concrete argument for a separate `IsTraining` column: the toggle should keep working for genuinely archived events and must never reach training ones.

### 3. Photos are exposed at the list, not the detail

`platform-event.interface.ts:79` — the interface comment says it plainly: *"Returned with list data when the event has image attachments."* Each image carries a `readUrl` SAS link. So a training event merely appearing in the list already hands live, time-limited image URLs to the browser, before anyone opens anything.

### 4. Action plans carry the guard's text permanently

`action-plan-modal.component.ts:189`:

```ts
copyEventDescriptionToRootCause(): void {
  const eventDescription = this.platformEventContext?.description?.trim();
  if (eventDescription) { this.model.rootCause = eventDescription; }
}
```

One button copies the incident description into the action plan's persisted `rootCause`. After that the text is ActionPlan data with no relationship to the event's flag, and no later filter reaches it. Block creation server-side for training events.

Credit where due: `platform-events-list.component.html:450` already hides the create button when `details.active` is false. That is the only place in the entire UI that checks the archive flag.

### 5. The SAM chat map is a separate backend

`chat.service.ts` defines `ChatStructuredLocationPoint` with `event_title`, `event_type` and `event_description`, rendered as clickable map pins with a detail panel. It is fed by `environment.apiUrl` — the conversational backend — **not** by `PlatformEventService`. It will not inherit any fix made to the PlatformEvent controller, and the chat PDF export rasterises the whole thing into a downloadable file.

## Good news

- **No event is linkable by URL.** The detail view is a modal, not a route; no code writes an event id to the URL. That contains the missing `Active` filter on `GetByIdAsync` to whatever reaches the list.
- **No CSV, Excel or print export** of event data anywhere.
- **No dashboard tile, chart or widget** counting events. `dashboard.service.ts` has no event endpoint.
- **The pagination count inherits** the filter automatically.

## Still unknown

**Power BI.** Embedded report pages are rendered from credentials in `PowerBIDetails`; the dataset is defined outside every repo. If any page queries `PlatformEvent`, it will not inherit anything done in code. Someone has to check this on the BI side.

## Still not found, across seven repos

NFC and long-press-to-edit. The frontend's fourteen `nfc` matches are Material Design Icon glyph names in a webfont; the seven `heatmap` matches are vendor chart libraries. Neither is a feature.

**The location heatmap does not exist in any repo read so far** — not backend, not frontend. The nearest thing is the SAM chat map at surface 4.

---

# ADDENDUM 3 — Pronect-Integrations, the missing bridge

This is the repo that connects SAM OnSite to the Pronect platform. Everything earlier in this document treated that link as an assumption. It is now read.

A small service: 19 C# files, one controller, one webhook handler.

## The path a guard's report actually takes

1. The Flutter app writes a row to the Supabase `events` table.
2. A Supabase dispatcher reads an outbox and POSTs to `webhooks/onsite/platform-event`, signed with an HMAC header.
3. `OnSiteMapper.Map()` turns the payload into a `PlatformEvent`.
4. Any photos are copied from Supabase storage into Azure Blob storage.
5. `PronectIngestionClient` POSTs to `/api/v1/PlatformEvent/IngestPlatformEvent` on the Pronect API with an AAD token.

Failures return 500 on purpose, so the Supabase outbox retries rather than marking the row done.

## The good news: the training flag is already plumbed

`OnSiteMapper.cs:65`:

```csharp
Active = TryGetBoolean(evt, "active")
    ?? TryGetBoolean(payload, "active")
    ?? true,
```

The mapper already reads `active` off the event, falls back to the envelope, and defaults to `true`. `PlatformEvent.Active` exists on the DTO, on the ingest DTO, on the entity, and the list query filters on it.

So a practice report can be marked at creation with **no change to this service and no change to the Pronect API**. The flag travels the whole way on its own.

## The bad news: one field is missing at the source

The Flutter `Event` model (`lib/data/models/event.dart`) has no `active` field. `toJson()` never emits one. So today every OnSite event arrives with `active` absent and ingests as `Active = true`.

That is the change: one field on the Dart model, one column on the Supabase `events` table, and the outbox payload carries it. Small, and at the right end — the flag is set where the report is born, which is what "marked as training at creation" requires.

## The defect this makes concrete

Setting `active = false` **does not stop the supervisor email.**

`PlatformEventService.IngestAsync` sends the critical-event mail on `if (isNewPlatformEvent)` alone. `Active` is never consulted. A practice report marked training from birth still wakes a supervisor.

Archive mode hides it from the list. It does not hold the alert.

## But there is a second lever, and it costs nothing

Recipients come from `GetEventNotificationRecipientsAsync(clientId, type, locationIds)`, and `NotificationScenarioRepository.cs:329`:

```csharp
if (matchingScenarios.Count == 0)
{
    return [];
}
```

No notification scenario matching the event's **type code** means no recipients, which means no email.

So a training event carrying its own `type` — one with no scenario configured — sends nothing, with no code change at all. Configuration only.

That gives two independent guards:

- **`active = false`** keeps it out of the lists, and is the durable flag every other consumer reads.
- **a distinct `type`** keeps the alert from firing.

Use both. Neither alone is sufficient: the type alone leaves the event visible in the list, and `active` alone leaves the alert firing.

## Three smaller things worth knowing

**Photos land in production storage before ingestion.** `StoreImagesAsync` runs before `IngestAsync` and is not conditional. A practice photo is copied into the `platform-event-images` container regardless of any flag. Cheapest fix for the MVP: no photo step in training.

**Re-ingest overwrites the flag.** Identity is `(SourceClientId, SourceSystem, SourceEventId)`, and on a repeat `existing.Active = entity.Active`. If a later webhook for the same event omits `active`, the mapper's `?? true` un-archives it. The flag must be written on the Supabase row itself, not passed once in an envelope.

**The bypass is environment-gated.** `x-dev-bypass` only works when the environment is `Localhost`, and the HMAC secret differs per environment. Nothing to fix.

## One thing to raise separately

`appsettings.Production.json` and `appsettings.Uat.json` carry live secrets in plaintext in the repo — AAD client secrets, storage account keys, and the webhook signing secrets for all environments.

This has nothing to do with onboarding, and I have not touched it. It should go to whoever owns the repo, because anyone with read access to it currently holds production storage keys.

## What this addendum changes in the plan

- "Marked as training at creation" moves from *unproven* to *confirmed feasible*, and gets cheaper — the platform end is already built.
- "Kept out of every live surface" keeps its cost, and gains a second required change: the alert path, which archive mode does not cover.
- The training event should carry its own `type`, not just the flag. That was not in the spec and should be.
