"""
field_schema.py  (TEST BUILD)

Single source of truth for every field on the agent intake dashboard.
Both app.py (renders the form + validates submissions) and the packet
generator rely on this list, so a field only has to be added/changed once.

Field dict keys:
    id          – form field name / data dict key (snake_case)
    label       – label shown to the agent
    type        – "text" | "textarea" | "number" | "select" | "yesno" |
                  "yesno_na" | "state"
    section     – "TOP" ("Preamble") or "BOTTOM" ("Contentions") — drives
                  the two-stage wizard order
    required    – bool, whether the field is mandatory
    help        – optional helper text shown under the field
    default_value – optional value the input is pre-filled with (agent can
                  edit or clear it)
    options     – for type="select"
    depends_on  – (field_id, "YES"|"NO") – field is only shown/enabled when
                  the referenced yes/no field has that value
    clause_preview – for type="yesno"/"yesno_na" fields that gate an
                  optional clause: the literal clause text shown to the
                  agent so they know what selecting "Yes" will add to the
                  document. Tokens are shown in {curly braces} as a hint of
                  what the agent will be asked to fill in.

"yesno_na" fields render a third "N/A" button. Selecting N/A (like "No")
never causes the associated clause to be inserted — it's a distinct value
from "No" purely for the agent's own record-keeping / clarity when a
question genuinely doesn't apply to the claim.
"""

from __future__ import annotations
from typing import Any, Dict, List, Optional

STATE_OPTIONS = [
    "AL", "AK", "AZ", "AR", "CA", "CO", "CT", "DE", "FL", "GA", "HI", "ID",
    "IL", "IN", "IA", "KS", "KY", "LA", "ME", "MD", "MA", "MI", "MN", "MS",
    "MO", "MT", "NE", "NV", "NH", "NJ", "NM", "NY", "NC", "ND", "OH", "OK",
    "OR", "PA", "RI", "SC", "SD", "TN", "TX", "UT", "VT", "VA", "WA", "WV",
    "WI", "WY", "DC",
]

# Display names shown to the agent (internal section keys stay "TOP"/"BOTTOM"
# to match TOP_Template.docx / BOTTOM_Template.docx).
SECTION_DISPLAY_NAMES = {"TOP": "Preamble", "BOTTOM": "Contentions"}

FIELDS: List[Dict[str, Any]] = [
    # ═══════════════════════════ PREAMBLE — CLAIM BASICS ═══════════════════
    {"id": "accident_state", "label": "Accident State", "type": "state",
     "section": "TOP", "required": True,
     "help": "Two-letter state where the accident occurred. Determines the Contentions case law inserted."},
    {"id": "repair_hours", "label": "Repair Hours", "type": "number",
     "section": "TOP", "required": True,
     "help": "Total billed shop hours. NOTE: Please refrain from using \"$\" or \"%\" symbols in your inputs for damage totals or liability %'s."},
    {"id": "property_damage", "label": "Property Damage Amount", "type": "number",
     "section": "TOP", "required": True,
     "help": "NOTE: Please refrain from using \"$\" or \"%\" symbols in your inputs for damage totals or liability %'s."},
    {"id": "towing", "label": "Towing Amount", "type": "number", "section": "TOP", "required": False},

    # ═══════════════════════════ PREAMBLE — NARRATIVE ═══════════════════════
    {"id": "accident_description", "label": "Accident Description", "type": "textarea",
     "section": "TOP", "required": True,
     "help": "NOTE: if no police report, reference the source of the narrative (driver's statement, accident report, etc.)"},

    {"id": "lane_direction_and_road_known", "label": "Are Lane/Direction/Road of Travel Known?", "type": "yesno",
     "section": "TOP", "required": True,
     "clause_preview": "Recovering Party was traveling in the {LANE_OF_TRAVEL} lane {DIRECTION_OF_TRAVEL} on {CLIENT_ROAD_OF_TRAVEL}. The Adverse Party was traveling in the {ADVERSE_LANE_OF_TRAVEL} lane {ADVERSE_DIRECTION_OF_TRAVEL} on {ADVERSE_ROAD_OF_TRAVEL}."},
    {"id": "lane_of_travel", "label": "Recovering Party Lane of Travel", "type": "text",
     "section": "TOP", "required": False, "depends_on": ("lane_direction_and_road_known", "YES")},
    {"id": "direction_of_travel", "label": "Recovering Party Direction of Travel", "type": "text",
     "section": "TOP", "required": False, "depends_on": ("lane_direction_and_road_known", "YES")},
    {"id": "client_road_of_travel", "label": "Recovering Party Road of Travel", "type": "text",
     "section": "TOP", "required": False, "depends_on": ("lane_direction_and_road_known", "YES")},
    {"id": "adverse_lane_of_travel", "label": "Adverse Party Lane of Travel", "type": "text",
     "section": "TOP", "required": False, "depends_on": ("lane_direction_and_road_known", "YES")},
    {"id": "adverse_direction_of_travel", "label": "Adverse Party Direction of Travel", "type": "text",
     "section": "TOP", "required": False, "depends_on": ("lane_direction_and_road_known", "YES")},
    {"id": "adverse_road_of_travel", "label": "Adverse Party Road of Travel", "type": "text",
     "section": "TOP", "required": False, "depends_on": ("lane_direction_and_road_known", "YES")},

    {"id": "collision_occurrence", "label": "Collision Occurrence", "type": "textarea",
     "section": "TOP", "required": True,
     "help": "One or two sentences describing how the impact happened. Use the Police Report/drivers statement for reference."},

    {"id": "police_report", "label": "Is There a Police Report?", "type": "yesno",
     "section": "TOP", "required": True,
     "clause_preview": "See narrative on page {PAGE_OF_PR_NARRATIVE} of the Police Report attached. See the scene diagram on page {PAGE_OF_DIAGRAM} of the Police Report for verification."},
    {"id": "page_of_pr_narrative", "label": "Page of PR Narrative", "type": "text",
     "section": "TOP", "required": False, "depends_on": ("police_report", "YES")},
    {"id": "page_of_diagram", "label": "Page of PR Scene Diagram", "type": "text",
     "section": "TOP", "required": False, "depends_on": ("police_report", "YES")},

    {"id": "video", "label": "Is There Video of the Accident?", "type": "yesno",
     "section": "TOP", "required": True,
     "clause_preview": "Please see the attached video for a visual account of the accident as it occurred."},

    # ── driver / supervisor / monitor statement (8-way combo clause) ──
    {"id": "drivers_statement", "label": "Driver's Statement Obtained?", "type": "yesno",
     "section": "TOP", "required": True,
     "clause_preview": "Recovering Party's driver also provided a signed written statement, consistent with the facts of loss and included in the Police Report."},
    {"id": "supervisor_statement", "label": "Supervisor's Statement Obtained?", "type": "yesno",
     "section": "TOP", "required": True,
     "clause_preview": "Recovering Party's supervisor also provided a signed written statement, consistent with the facts of loss and included in the Police Report."},
    {"id": "monitor_statement", "label": "Monitor's Statement Obtained?", "type": "yesno",
     "section": "TOP", "required": True,
     "clause_preview": "Recovering Party's monitor also provided a signed written statement, consistent with the facts of loss and included in the Police Report."},

    # ── negligence theory bullets ──
    {"id": "adverse_contributing_action_2", "label": "Adverse Contributing Action #2", "type": "text",
     "section": "TOP", "required": True,
     "default_value": "Failure to Maintain a Proper Lookout",
     "help": "Listed under NEGLIGENCE THEORY, alongside \"Driver Inattention/Distraction\". Pre-filled — edit or clear as needed for this claim."},
    {"id": "adverse_contributing_action_3", "label": "Adverse Contributing Action #3 (optional)", "type": "text",
     "section": "TOP", "required": False,
     "help": "Leave blank to omit this bullet. Also used in the Favorable PR clause below if you select Yes to that."},
    {"id": "adverse_contributing_action_4", "label": "Adverse Contributing Action #4 (optional)", "type": "text",
     "section": "TOP", "required": False, "help": "Leave blank to omit this bullet entirely."},

    # ── PR contributing factors ──
    {"id": "contributing_factor_on_pr_adverse", "label": "Contributing Factor Cited Against Adverse on PR?", "type": "yesno_na",
     "section": "TOP", "required": True,
     "clause_preview": "On page {PR_PAGE_CONTAINING_CONTRIBUTING_FACTORS_ADVERSE} of the attached Police Report, the Investigating Officer cited Adverse Party with {PR_CONTRIBUTING_ACTION_ADVERSE}."},
    {"id": "pr_page_containing_contributing_factors_adverse", "label": "PR Page — Adverse Contributing Factors", "type": "text",
     "section": "TOP", "required": False, "depends_on": ("contributing_factor_on_pr_adverse", "YES")},
    {"id": "pr_contributing_action_adverse", "label": "PR Contributing Action Cited (Adverse)", "type": "text",
     "section": "TOP", "required": False, "depends_on": ("contributing_factor_on_pr_adverse", "YES")},

    {"id": "contributing_factor_on_pr_client", "label": "Contributing Factor Cited Against Recovering Party (Client) on PR?", "type": "yesno_na",
     "section": "TOP", "required": True,
     "clause_preview": "(Shown only if you answer \"No\") — On page {PR_PAGE_CONTAINING_CONTRIBUTING_FACTORS_CLIENT} of the attached Police Report, the Investigating Officer cited Recovering Party with no contributing factors."},
    {"id": "pr_page_containing_contributing_factors_client", "label": "PR Page — Client Contributing Factors", "type": "text",
     "section": "TOP", "required": False, "depends_on": ("contributing_factor_on_pr_client", "NO")},

    {"id": "adverse_issued_citation", "label": "Was Adverse Party Issued a Citation?", "type": "yesno_na",
     "section": "TOP", "required": True,
     "clause_preview": "The Adverse Party was in violation of {STATE_VTL} traffic law, {TRAFFIC_LAW}, which states: {TRAFFIC_LAW_DESCRIPTION}."},
    {"id": "state_vtl", "label": "State Traffic Law Code (e.g. NY VTL)", "type": "text",
     "section": "TOP", "required": False, "depends_on": ("adverse_issued_citation", "YES")},
    {"id": "traffic_law", "label": "Traffic Law / Statute Cited", "type": "text",
     "section": "TOP", "required": False, "depends_on": ("adverse_issued_citation", "YES")},
    {"id": "traffic_law_description", "label": "Traffic Law Description", "type": "textarea",
     "section": "TOP", "required": False, "depends_on": ("adverse_issued_citation", "YES")},

    {"id": "pr_page_showing_adverse_point_of_impact", "label": "PR Page — Adverse Point of Impact", "type": "text",
     "section": "TOP", "required": False, "depends_on": ("police_report", "YES")},
    {"id": "adverse_point_of_impact", "label": "Adverse Party Point of Impact", "type": "text", "section": "TOP", "required": True},
    {"id": "pr_page_showing_client_point_of_impact", "label": "PR Page — Client Point of Impact", "type": "text",
     "section": "TOP", "required": False, "depends_on": ("police_report", "YES")},
    {"id": "client_point_of_impact", "label": "Recovering Party Point of Impact", "type": "text", "section": "TOP", "required": True},

    {"id": "photos_of_damages", "label": "Photos of Damages Available?", "type": "yesno",
     "section": "TOP", "required": True,
     "clause_preview": "The points of impact and photos (and video, if applicable) prove that Recovering Party was there to be seen, and Adverse Party had the last opportunity to avoid this collision."},
    {"id": "liability_obvious", "label": "Is Liability Obvious?", "type": "yesno",
     "section": "TOP", "required": True,
     "clause_preview": "The points of impact and photos (and video, if applicable) also prove that the Adverse Party's actions left the Recovering Party without the opportunity to take evasive action and that Adverse Party was the sole cause of this collision."},
    {"id": "favorable_pr", "label": "Favorable Police Report?", "type": "yesno",
     "section": "TOP", "required": True,
     "clause_preview": "Police investigated and narrated their findings in a formal report... The report was clear and cited the Adverse Party's {ADVERSE_CONTRIBUTING_ACTION_3} as the cause of this loss."},

    # ── total loss taxi ──
    {"id": "total_loss_taxi", "label": "Is this a Taxi Total-Loss Claim?", "type": "yesno",
     "section": "TOP", "required": True,
     "clause_preview": "When a taxi is deemed a total loss, additional expenses are incurred... The amount required for the equipment transfer is ${TAXI_EQUIPMENT_TRANSFER_COST}. Please see \u201cmake ready\u201d fee sheet attached."},
    {"id": "taxi_equipment_transfer_cost", "label": "Taxi Equipment Transfer Cost", "type": "number",
     "section": "TOP", "required": False, "depends_on": ("total_loss_taxi", "YES")},

    # ── PD payment status (shared TOP/BOTTOM) ──
    {"id": "adverse_paid_100_pct_pd", "label": "Adverse Party Paid 100% of Property Damage?", "type": "yesno",
     "section": "TOP", "required": True,
     "clause_preview": "Adverse Party refunded Recovering Party 100% of the ${PROPERTY_DAMAGE_AMOUNT}. See check attached in Feature Information section."},
    {"id": "pd_amount_agreed_upon", "label": "Was a Property-Damage Amount Agreed Upon?", "type": "yesno",
     "section": "TOP", "required": True,
     "clause_preview": "Adverse Party refunded the Recovering Party ${PROPERTY_DAMAGE_AMOUNT_AGREED_UPON}. The Adverse Party and the Recovering Party agreed on this amount to satisfy the physical repairs."},
    {"id": "property_damage_amount_agreed_upon", "label": "Agreed-Upon Property Damage Amount", "type": "number",
     "section": "TOP", "required": False, "depends_on": ("pd_amount_agreed_upon", "YES")},
    {"id": "adverse_paid_partial_pd", "label": "Adverse Party Paid Partial Property Damage?", "type": "yesno",
     "section": "TOP", "required": True,
     "clause_preview": "Adverse Party partially refunded Recovering Party for property damage in the amount of ${PARTIAL_PD_AMOUNT}... The amount Adverse Party refunded was {PERCENTAGE_OF_ADVERSE_ESTIMATE_RECOVERED}% of Adverse Party's desk review."},
    {"id": "partial_pd_amount", "label": "Partial Property Damage Amount Paid", "type": "number",
     "section": "TOP", "required": False, "depends_on": ("adverse_paid_partial_pd", "YES")},
    {"id": "percentage_of_adverse_estimate_recovered", "label": "% of Adverse Estimate Recovered", "type": "number",
     "section": "TOP", "required": False, "depends_on": ("adverse_paid_partial_pd", "YES"),
     "help": "Enter as a plain number, e.g. 60 (do not include a % symbol)."},

    {"id": "liability_contested", "label": "Is Liability Contested?", "type": "yesno",
     "section": "TOP", "required": True,
     "clause_preview": "(Shown only if you answer \"No\") — Liability is not an issue."},

    # ── loss of use total (typed by the agent, not calculated) ──
    {"id": "total_lou_days", "label": "Total Loss-of-Use Days", "type": "number",
     "section": "TOP", "required": True,
     "help": "Used for both the Preamble ([TOTAL_LOU_DAYS]) and Contentions ([TOTAL_DAYS]) documents — enter once here."},
    {"id": "loss_of_use_amount", "label": "Loss of Use Amount", "type": "number",
     "section": "TOP", "required": True,
     "help": "Used for both documents, and included in the Grand Total."},

    # ── loss of use payment status ──
    {"id": "partial_lou", "label": "Did Adverse Party Partially Pay Loss of Use?", "type": "yesno",
     "section": "TOP", "required": True,
     "clause_preview": "The Adverse Party partially refunded the Recovering Party for loss of use in the amount of {PARTIAL_LOSS_OF_USE_PAYMENT}."},
    {"id": "partial_loss_of_use_payment", "label": "Partial Loss of Use Payment Amount", "type": "number",
     "section": "TOP", "required": False, "depends_on": ("partial_lou", "YES")},
    {"id": "partial_payment_combo_check", "label": "Was the Partial Payment a Combined PD + LOU Check?", "type": "yesno",
     "section": "TOP", "required": True,
     "clause_preview": "Adverse Party check is in the amount of {TOTAL_ADVERSE_PARTIAL} which included the {TOTAL_ADVERSE_PD_PAYMENT} for property damage and {PARTIAL_LOSS_OF_USE_PAYMENT} for loss of use damage."},
    {"id": "total_adverse_partial", "label": "Total Combined Check Amount", "type": "number",
     "section": "TOP", "required": False, "depends_on": ("partial_payment_combo_check", "YES")},
    {"id": "total_adverse_pd_payment", "label": "Portion of Check for Property Damage", "type": "number",
     "section": "TOP", "required": False, "depends_on": ("partial_payment_combo_check", "YES")},
    {"id": "adverse_paid_partial_lou_without_supports", "label": "Did Adverse Pay a Partial LOU Amount Without Support?", "type": "yesno",
     "section": "TOP", "required": True,
     "clause_preview": "If Adverse Party wishes to pursue a lesser amount for loss of use, they must support the reduction through their own independent evidence and logic. Simply refunding an arbitrary amount does not exempt Adverse Party from their burden of proof. The Recovering Party has made good-faith efforts to resolve this discrepancy without resorting to arbitration..."},
    {"id": "adverse_offered_partial_lou_without_supports", "label": "Did Adverse Offer (but Not Pay) a Partial LOU Amount Without Support?", "type": "yesno",
     "section": "TOP", "required": True,
     "help": "For an offer that was made but not accepted/paid — distinct from the field above, which is for an amount Adverse actually paid.",
     "clause_preview": "The amount Adverse Party offered for loss of use is an arbitrary amount not based on any reasonable rental rate for a like-in-kind vehicle."},
    {"id": "previous_vendor_lou_demand", "label": "Was a Previous LOU Demand Sent by a Prior Vendor/TPA?", "type": "yesno",
     "section": "TOP", "required": True,
     "clause_preview": "A demand for loss of use was originally provided to Adverse Party in {YEAR_PREVIOUS_DEMAND_SUBMITTED}... In {YEAR_DTS_DEMAND_SENT}, Downtime Subrogation provided an updated and corrected demand..."},
    {"id": "year_previous_demand_submitted", "label": "Year Previous Demand Submitted", "type": "text",
     "section": "TOP", "required": False, "depends_on": ("previous_vendor_lou_demand", "YES")},
    {"id": "year_dts_demand_sent", "label": "Year DTS Demand Sent", "type": "text",
     "section": "TOP", "required": False, "depends_on": ("previous_vendor_lou_demand", "YES")},

    # ── diminished value ──
    {"id": "pursuing_diminished_value", "label": "Pursuing Diminished Value?", "type": "yesno",
     "section": "TOP", "required": True,
     "clause_preview": "Recovering Party also suffered ${DIMINISHED_VALUE} in diminutive value... The Restatement of Torts, Second \u00a7 928 also affirms that a plaintiff may be reimbursed for the difference in value of his or her damaged vehicle before and after repair."},
    {"id": "diminished_value_amount", "label": "Diminished Value Amount", "type": "number",
     "section": "TOP", "required": False, "depends_on": ("pursuing_diminished_value", "YES"),
     "help": "Included in the Grand Total."},

    # ═══════════════════════════ CONTENTIONS — LOU / DAMAGES ═══════════════
    # ── these were previously auto-calculated from the demand letter
    #    generator's state/vehicle rate tables. For this test build, the
    #    agent enters each value directly. ──
    {"id": "repl_daily", "label": "Daily Replacement/Rental Value ($/day)", "type": "number",
     "section": "BOTTOM", "required": True,
     "help": "The published daily rental rate for a like-in-kind vehicle, including a driver if applicable."},
    {"id": "driver_wage_hourly", "label": "Average Local Driver Wage (hourly)", "type": "number",
     "section": "BOTTOM", "required": True},
    {"id": "driver_wage_daily", "label": "Average Local Driver Wage (daily = hourly × 8)", "type": "number",
     "section": "BOTTOM", "required": True,
     "help": "Enter the hourly wage above multiplied by 8."},
    {"id": "final_daily_rate", "label": "Final Daily Rate (after driver-wage deduction, if applicable)", "type": "number",
     "section": "BOTTOM", "required": True,
     "help": "Typically Daily Replacement Value minus Average Daily Driver Wage."},
    {"id": "raw_days", "label": "Raw Repair Days (before weekends/paint-cure day)", "type": "number",
     "section": "BOTTOM", "required": True},
    {"id": "applicable_weekend_days", "label": "Applicable Weekend Days", "type": "number",
     "section": "BOTTOM", "required": True,
     "help": "Number of weekend days falling within the repair period, added on top of the raw repair days."},

    {"id": "driver_salary_reduction", "label": "Deducting Driver's Salary From the LOU Demand?", "type": "yesno",
     "section": "BOTTOM", "required": True,
     "clause_preview": "Controls whether the LOU rate sentence quotes the full daily rate or the rate after the driver-wage deduction, and adds: \u201cWE ARE NOT SEEKING REIMBURSEMENT FOR DRIVER\u2019S SALARY...\u201d"},
    {"id": "adverse_states_no_lou_support_but_paid_full_pd", "label": "Did Adverse State No LOU Support Was Provided (but Paid Full PD)?", "type": "yesno",
     "section": "BOTTOM", "required": True,
     "clause_preview": "Adverse Party has asserted that we have not adequately supported our loss of use claim... Adverse party cannot unilaterally refuse to pay a loss of use claim which has been lawfully supported."},
    {"id": "attaching_previous_lou_arb_decisions", "label": "Attaching Previous LOU Arbitration Decisions?", "type": "yesno",
     "section": "BOTTOM", "required": True,
     "clause_preview": "The issue of loss of use in [state] has been previously decided by Arbitration Forums. See copy attached."},
    {"id": "adverse_insurance_carrier_is_progressive", "label": "Is the Adverse Insurance Carrier Progressive?", "type": "yesno",
     "section": "BOTTOM", "required": True,
     "clause_preview": "Adds the full Article First self-insured eligibility argument, including the state's progressive-specific case law citation."},
]

TOP_FIELDS = [f for f in FIELDS if f["section"] == "TOP"]
BOTTOM_FIELDS = [f for f in FIELDS if f["section"] == "BOTTOM"]


def field_by_id(field_id: str) -> Optional[Dict[str, Any]]:
    for f in FIELDS:
        if f["id"] == field_id:
            return f
    return None
