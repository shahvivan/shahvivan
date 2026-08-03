**To:** Jeroen
**Cc:** Vincent Smeyers
**Subject:** Re: SAM OnSite guard onboarding — updated specification
**Attachments:** SAM_OnSite_Guard_Onboarding_Spec.docx · SAM_OnSite_Guard_Onboarding_Detailed_Spec.docx

---

Hi Jeroen,

Thanks for the comments — they were the right ones and all of them are in. Two documents attached:

- **Guard Onboarding — specification** (5 pages): the whole feature in one read, for you and Vincent.
- **Detailed specification** (20 pages): requirements, acceptance criteria, test matrix and traceability, for engineering and QA.

**Your points**

*Device setup.* The three gates are the flow now, in your order: permissions, then NFC, then background tracking. Each duplicates the link into onboarding rather than sending the guard to the app's settings screen, and each returns to the flow. Then the scenarios.

*Tips and tricks.* Found the General recommendations slide. All eight tips are mapped to the step where each becomes useful, and two of them are not taught at all — they are done. Press-and-hold to fix a spelling mistake *is* the manual-correction step, and speaking in full sentences *is* what the report prompt asks for. The rest arrive as a short aside from SAM at the moment they apply. The tip-by-tip mapping is in the detailed document.

*Replacing the PowerPoint steps.* The flow covers the app-setup slides: background tracking, permissions, NFC, and their screenshot walkthroughs. Installation and first log-in sit on the same slide but cannot move into the app, since the flow only exists once the guard has signed in. Those two stay with the trainer.

*Engagement.* Agreed — and your own slide put it best: "SAM is your new personal assistant, so learn to know each other." The opening uses that framing, and a second screen tells the guard SAM can be taught the words they use, so it reads as theirs from the first screen. Tone is a hard requirement in the spec rather than guidance: a dry "do this, then that" sequence is an explicit fail condition, tested with guards.

*Urgent incident.* Cut right back, as you suggested.

**Three things I have done differently, and why**

1. **The urgent-incident path is now required rather than optional** — but for a different reason than incident reporting. Walking the flow as a guard, it turns out to be the only way out of something that cannot be skipped. If a guard gets stuck on a step, that is the only door. Same small mechanism, load-bearing for a different purpose.

2. **Technically checking the settings is in the MVP, for permissions.** You suggested it as a later maybe. The app's Location screen already reads permission state, so the mechanism exists — the permissions gate now reads back rather than asks, and a permission still off returns the guard to that gate. NFC and background tracking stay as the guard confirming, because I cannot tell whether their state is readable. That is my question below.

3. **Device readiness is now part of the definition of done.** Your line "at that point all settings are activated" made me notice the spec never said that — it defined done as the three actions alone. A guard who finishes and then cannot scan a checkpoint has not been onboarded.

**One question back**

Permission state is readable today. Can NFC and background-tracking state be read the same way? If so, all three gates verify rather than ask, and the promise that the phone is ready when onboarding ends is fully kept. If not, those two stay on the guard's word and I have written the residual risk into the error table.

**One thing worth flagging**

The specification was careful about practice data reaching live operations and silent about the reverse. A guard thirty seconds into being taught "this is how you report things" has something actually happen, presses Talk and reports it — and the training flag hides it from the supervisor it should have reached. That is now a requirement. It does not go away at go-live either, because the refresher puts experienced guards back in the same position.

Also reflecting Vincent's note: the filter that keeps training data out of live surfaces is not built, so go-live is blocked on it and this runs in DEV/UAT until it exists.

On tooling — I used AI to produce the documents and to run the checks behind them, including an independent review and a walkthrough of the flow as a guard would experience it, which is what surfaced the point above. The content, the decisions and the scope calls are mine.

Happy to walk through either document.

Best,
Vivan
