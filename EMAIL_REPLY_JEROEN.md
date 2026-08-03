**To:** Jeroen
**Cc:** Vincent Smeyers
**Subject:** Re: SAM OnSite guard onboarding — updated specification
**Attachments:** SAM_OnSite_Guard_Onboarding_Spec.docx · SAM_OnSite_Guard_Onboarding_Detailed_Spec.docx

---

Hi Jeroen,

Thanks, useful feedback. Both documents attached: a short one for you and Vincent, a longer one with the requirements and tests for engineering.

It's all in. Two things I did differently, which I'd rather flag than have you find. The urgent-incident exit stays, because walking the flow as a guard it's the only way out of something that can't be skipped. And device readiness is part of "done" now, which the spec never actually said.

One question, and it's the one you raised. Right now all three gates work by asking the guard to confirm they've switched something on, which means someone can tap Done without having done it and nobody finds out until the app fails on them weeks later. Can the app read whether any of these are actually on, so it checks instead of asks? I couldn't establish that from anything I had. Whichever of the three turn out to be checkable should be in the build.

Worth a look. The spec guarded against practice data reaching live operations, but not the reverse. A guard half a minute into "this is how you report things" has something real happen, presses Talk, and the training flag hides it from their supervisor. That's a requirement now, and the refresher keeps it live after go-live.

I used AI for the documents and the reviews behind them. The calls are mine.

Happy to talk it through.

Vivan
