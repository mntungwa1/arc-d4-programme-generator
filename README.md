# ARC D4 Programme Generator

This Streamlit app turns the ARC D4 Simplified Design into a guided workflow. It reads the included Automation Matrix as its source of truth and lets users upload an amended workbook when the programme configuration changes.

## Run locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

## What it does

- Shows the 19 stages in their required five-lane order, including each entry condition, exit condition, gate and executable action.
- Lets the analyst select an innovation from the governed portfolio and complete its Investment-Ready Innovation Profile (IRIP).
- Displays the associated field dictionary, research tasks, gap linkages and five-pillar response.
- Enforces the four refusal conditions and the twelve IRIP component check before allowing progression.
- Exports the selected profile, audit log, pathway and validation rules as a controlled working record.

The app deliberately does not create plausible content for missing evidence. It identifies a gap, directs the user to the relevant research process, and records the completion decision.
