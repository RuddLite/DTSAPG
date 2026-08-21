# Arbitration Packet Generator — STANDALONE TEST BUILD

This is a **test-only variant** of the Arbitration Packet Generator, meant
for you to deploy on your own Render instance separately from the
production Herd build.

**The only difference from the production version:** this build does not
depend on `demand_letter_generator.py` at all (the file isn't even
included). Every value that the production version calculates from the
demand letter generator's state/vehicle rate tables is instead typed in
directly by the agent:

| Token | Production version | This test build |
|---|---|---|
| `REPL_DAILY` | looked up from `STATE_VEHICLES[state][vehicle]` | agent types it |
| `DRIVER_WAGE_HOURLY` | looked up from `DRIVER_WAGE[state]` | agent types it |
| `DRIVER_WAGE_DAILY` | `hourly wage × 8` | agent types it |
| `FINAL_DAILY_RATE` | `REPL_DAILY − DRIVER_WAGE_DAILY` | agent types it |
| `RAW_DAYS` | `ceil(hours / 4)` | agent types it |
| `TOTAL_DAYS` / `TOTAL_LOU_DAYS` | raw days + weekends + 1 paint-cure day | agent types it once (used in both docs) |
| `LOSS_OF_USE_AMOUNT` | `TOTAL_DAYS × FINAL_DAILY_RATE` | agent types it (used in both docs) |
| `DIMINISHED_VALUE` | flat 10% of property damage | agent types it |

The `Vehicle Type` and `Date of Loss` dashboard fields are also removed in
this build — they existed only to feed the rate-table lookup and don't
appear as tokens in either template, so there's nothing left for them to
drive here. (Date of Loss can easily be re-added as a plain reference field
if you want it captured for your own records — it just won't affect any
calculation in this build.)

Everything else — conditional clauses, the optional negligence bullets,
state case-law insertion, the two-step wizard, clause previews, dual
document download — is identical to the production version. See the main
`README.md` pattern from the production build for how the merge engine
works; the only file that changed materially is `arb_packet_generator.py`
(no `demand_letter_generator` import, `compute_lou()` removed) and
`field_schema.py` (added manual fields for the values above, removed
`vehicle_type`/`date_of_loss`).

## Running locally / deploying to Render

```
pip install -r requirements.txt
python app.py
# → http://127.0.0.1:5000  (Basic Auth: APP_USER / APP_PASS env vars, default cash/staff)
```

Render: point it at this folder (or its own repo) with `web: gunicorn
app:app` (already in `Procfile`) and `requirements.txt`. Set `APP_USER` /
`APP_PASS` env vars in the Render dashboard for your own login.

## CLI

```
python arb_packet_generator.py --json sample_intake.json --out output/
```

## Case law

Same as production: only NY and TN are populated in
`templates/Case_Law_for_BOTTOM.docx`. Append more states there in the same
`[XX_CASE_LAW]:` / `[XX_PROGRESSIVE_CASE_LAW]:` format — no code changes
needed.
