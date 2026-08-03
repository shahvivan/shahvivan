**To:** Jeroen
**Cc:** Vincent Smeyers
**Subject:** Re: SAM OnSite guard onboarding — updated specification
**Attachments:** SAM_OnSite_Guard_Onboarding_Spec.docx · SAM_OnSite_Guard_Onboarding_Detailed_Spec.docx

---

Hi Jeroen,

All your points are in. Two documents attached: a 5-page brief for you and Vincent, and a detailed version for engineering and QA.

Three places I went further or differently than you suggested:

- **Permissions are now checked, not asked.** The Location screen already reads permission state, so the gate reads back and returns the guard if something is off. NFC and background tracking still ask — see my question below.
- **The urgent-incident path is required rather than optional**, for a different reason than incident reporting: walking the flow as a guard, it is the only way out of something that cannot be skipped.
- **Device readiness is part of the definition of done.** Your "at that point all settings are activated" made me notice the spec never said so — it defined done as the three actions alone.

On the tips: all eight from the General recommendations slide are placed where each becomes useful, and two are not taught at all. Press-and-hold to fix a spelling mistake *is* the manual-correction step.

**My question:** permission state is readable today. Can NFC and background-tracking state be read the same way? If so, all three gates verify rather than ask.

**Worth your attention:** the spec guarded against practice data reaching live operations, but not the reverse — a guard reporting something real during practice, then hidden from their supervisor by the training flag. That is now a requirement, and it persists after go-live through the refresher.

I used AI to produce the documents and run the reviews behind them, including a walkthrough of the flow as a guard would experience it, which is what surfaced that last point. The decisions are mine.

Happy to walk through either document.

Best,
Vivan
