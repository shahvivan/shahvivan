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
