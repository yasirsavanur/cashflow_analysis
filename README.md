# MegaWidget Cashflow Analysis

<i>Tools used: Excel with VBA automation, Pivot Tables, and Scenario Analysis</i>

Welcome to my quick walkthrough of how I explored and broke down the cash flow dynamics for Megawidget International. This report is more about *practical insights* and *what worked*—not heavy jargon.

---

## 1. What’s This All About?

The goal? Build an interactive spreadsheet that could help a mid-size manufacturing company—Megawidget International—understand its cash flow better. Especially as they plan to invest in a new manufacturing facility for their high-performing products (B, D, and E).

The workbook helps answer questions like:
- Can the company afford to take a loan for expansion?
- How much should they really be spending on ads?
- What happens to profits if we tweak margins or discount offers?

---

## 2. Clean-Up Crew: Organising the Raw Data

**Sheet:** `1. Clean`

I used a macro to tidy up the raw sales data:
- Removed blanks and irrelevant rows
- Formatted dates consistently
- Made sure all entries matched expected product and region codes

This step sets the foundation for every single visual or metric later in the workbook.

---

## 3. Sales Summary Snapshot

**Sheet:** `2. Organise`

Two macros keep this updated:
- One refreshes the cleaned data
- The other recalculates regional and product-based sales summaries

### Findings:
- **Gigastuff (Product B)** is super stable — always hovering near 2000 units/month.
- **Hyperwotsit (Product D)** is the most unpredictable — big highs and lows.
- **Megawidget S (A1)** is showing consistent growth across regions like WC and WS.
- On the flip side, **Megawidget L (A3)** is fading out — sales dropped every quarter.

---

## 4. Forecasting the Cash Flow for 2023

**Sheets:** `3.1 Cash Flow Forecast` and `3.2 Cash Flow Dashboard`

I played around with different loan terms, ad budgets, and margin assumptions.

### Default Scenario:
- Loan: £650,000 over 10 years @ 4.3% interest
- Profit Margin: 25%
- Ads: £30,000/month

**Result:** The company would lose money in 10 out of 12 months. Not ideal.

### Optimised Scenario:
- Loan extended to 15 years
- Margin bumped to 28%
- Ads trimmed to £20,000/month

**Result:** Year-end profit hits £157,410. Huge difference.

---

## 5. Switching Delivery Partners?

**New offer:** Flat £14.95 per order

Ran a quick analysis across 3 years of past sales.

**Finding:** They would’ve saved up to **46%** in delivery costs in 2022 alone. That’s a no-brainer.

---

## 6. Discounts: Helpful or Hurtful?

Simulated offering:
- 3% off for 20+ cases
- 5% for 50+
- 7% for 100+

With current sales trends, this plan would shrink:
- Revenue by ~4.7%
- Profits by ~6.8%

Unless you're expecting massive order volume boosts, this might hurt more than help.

---

## 7. Predicting Sales Growth (Two Ways)

### A. Deterministic Model
- Assumes steady 5% growth every month in 2023
- Starts from Dec 2022 sales

Result: Gradual and safe rise in monthly revenue.

### B. Monte Carlo Simulation
- Used random changes based on historical variance
- Ran 1000 iterations

Outcome:
- **Best case:** Revenue shoots up to £200,000
- **Worst case:** Dips below £50,000
- **Most likely (average):** £87,000/month

---

## Final Thoughts

This model doesn’t just spit out numbers—it tells a story. One where Megawidget International needs to:
- Rethink its finance plan if it wants to stay profitable
- Be cautious with discount strategies
- Seriously consider that delivery deal
- Stay optimistic, but plan for volatility in sales

The best part? Most of this analysis runs with just a click or two, thanks to VBA automation and pivot magic.

---

<i>Hope this gives a clear, honest look into how I approached cash flow analysis using Excel—and had a bit of fun with it too.</i>
