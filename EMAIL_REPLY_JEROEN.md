**To:** Jeroen
**Cc:** Vincent Smeyers
**Subject:** Re: SAM OnSite guard onboarding — updated specification
**Attachments:** SAM_OnSite_Guard_Onboarding_Spec.docx · SAM_OnSite_Guard_Onboarding_Detailed_Spec.docx

---

Hi Jeroen,

Thanks, useful feedback. Both documents attached: a short one for you and Vincent, a longer one with the requirements and tests for engineering.

Everything you asked for is in. Three things I did differently, which I'd rather flag than have you find.

Permissions are checked now, not asked. You had this down as a later maybe, but the Location screen already reads permission state, so the mechanism exists. The gate reads back and returns the guard if something is still off. NFC and background tracking still just ask, which is my question below.

The urgent-incident exit went the other way. You're right that it doesn't matter much for reporting incidents. But walking the flow as a guard on a first shift, it turned out to be the only way out of something that can't be skipped. So it stays, for a different reason than it was there before.

Your line about all the settings being active at that point made me realise the spec never actually said so. It defined done as the three actions and nothing about the phone. A guard who finishes and then can't scan a checkpoint hasn't been onboarded. That's in the definition of done now.

On the tips, all eight are placed where each one becomes useful. Two of them aren't taught at all, which I liked: press-and-hold to fix a spelling mistake is the manual correction step, so the guard just does it.

My question: permission state is readable. Can NFC and background tracking be read the same way? If so, all three gates verify instead of asking.

One thing I'd want you to look at. The spec was careful about practice data reaching live operations and said nothing about the reverse. A guard is half a minute into being told "this is how you report things", something real happens, they press Talk, and the training flag hides it from the supervisor it should have gone to. That's a requirement now. It doesn't go away at go-live either, because the refresher puts experienced guards back in the same spot.

I used AI for the documents and the reviews behind them, including a walkthrough of the flow from a guard's point of view, which is what caught that last one. The calls are mine.

Happy to go through either of them.

Vivan
