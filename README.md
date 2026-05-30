# RRSP Multi-Year Deduction Optimizer

An automated, data-driven financial modeling tool designed to maximize the Net Present Value (NPV) of Canadian Registered Retirement Savings Plan (RRSP) tax refunds. 
Built entirely within Excel using a heuristic optimization sweep (Hill-Climbing Local Search Algorithm), this VBA macro determines the mathematically ideal allocation of 
multi-year RRSP deductions across fluctuating income tax brackets. 

This algorithm works because TVM is a strictly smooth, monotonic exponential decay curve and the Canadian tax schedule is a non-decreasing step-function.
When you combine an upward-trending step function with a smooth exponential decay curve, you get a well-behaved, single-peaked landscape. Hence the heurisitc optimzation will find the global maximum of the landscape.

---

## 📊 Theoretical Background & Relevance

### The Canadian Tax Structure

Canada employs a **progressive tax system** where your income is taxed in tranches or "brackets." As your income increases, each dollar earned above a specific threshold is taxed at a higher rate (your Marginal Tax Rate). 
In Ontario, for example, marginal tax rates can step up significantly—ranging from around 20% on lower income brackets to over 53% on income exceeding $250,000.

### The Power of RRSP Deduction Arbitrage

An RRSP is not simply a tax-sheltered investment account that also serves as a tax arbitrage tool. 
Contributions made to an RRSP are deductible from your taxable income, effectively yielding a tax refund proportional to your marginal tax rate for that year.

However, Canadian tax law contains a highly valuable feature: **you do not have to claim the tax deduction in the same year you make the contribution.** You can legally carry forward undeducted contributions indefinitely.

This introduces a complex multi-variable optimization problem:

* **Option A:** Claim the deduction immediately to get cash back today, allowing that refund to compound over time (Time Value of Money).
* **Option B:** Defer claiming the deduction to a future year when your salary is projected to rise into a higher marginal tax bracket, yielding a larger absolute refund.

### Project Relevance

Most generic retirement calculators assume a flat tax rate or require manual trial-and-error. 
This project automates the decision-making process. 
By evaluating your projected salary trajectory against official tax brackets and an assumed investment discount rate, the algorithm  
considers both the **Time Value of Money** and **Tax Bracket Arbitrage** to secure the highest possible lifetime wealth.

---

## ⚖️ Strategic Framework: RRSP vs. TFSA Priority

To get the most out of this tool, it must be paired with the right cash allocation. 
A common misconception is that you should always max out one account before the other. 
In reality, the decision is dictated by comparing your **current marginal tax bracket** against your **expected marginal tax bracket in retirement**.

### The Golden Rule

* **Prioritize the RRSP** if your marginal tax rate **now** is *higher* than your expected marginal tax rate in **retirement** (Positive Tax Arbitrage).
* **Prioritize the TFSA** if your marginal tax rate **now** is *lower* or *equal* to your expected marginal tax rate in **retirement** (Avoids Negative Tax Arbitrage).

---

## 🛠️ Technical Implementation & Architecture

The project consists of an interactive Excel workspace backed by an automated optimization routine written in VBA.

### File Structure

* **`RRSP Deduction Model.xlsm`**: The macro-enabled workbook housing the user inputs, provincial/federal tax bracket formulas, and financial summary cards.
* **Column M (`RRSP Deduction Claimed`)**: The core decision variable manipulated by the optimization engine.
* **Column N (`RRSP Unclaimed Deductions`)**: The dynamic state tracking variable that rolls over carried-forward pools to subsequent years using the formula: `=N[t-1] + K[t] - M[t]`.
* **Cell Q20 (`Total PV Refund`)**: The objective function target cell calculating the Net Present Value of all refunds over the time horizon.



### The Optimization Algorithm 

Because Excel's native Solver can be unstable when handling complex multi-year piecewise tax equations across a long horizon, this model utilizes a custom native VBA heuristic loop.

The macro functions as follows:

1. **Establishes a Baseline:** Evaluates the spreadsheet’s current state and records the initial `Total PV Refund`.
2. **Executes a Time-Shifting Sweep:** Loops through every historical year $i$ and compares it to every future year $j$.
3. **Simulates Changes:** Shaves a small block of deductions ($\$500$) from year $i$ (where income might be lower) and reallocates it to year $j$ (where income might be higher).
4. **Evaluates Objective Function:** Checks cell `S6`. If the global NPV increases, the change is permanently locked in. If it decreases or stays the same, the shift is rolled back.
5. **Convergence:** Continues iterating through the entire multi-year matrix until a full pass results in absolutely zero incremental improvements to the total PV.

---

## Getting Started

### Prerequisites

* Microsoft Excel (Desktop Version recommended for macro performance).
* Macros must be enabled upon opening the workbook.

### Instructions

1. Open `RRSP Deduction Model.xlsm`.
2. Enter the desired number of years to project in "Total Years".
3. Click on the "Resize Table to Time Horizon" button.
4. Populate your **Projected Salary** profile and desired annual **RRSP Contribution Amounts** based on your savings capacity.
5. Ensure your discount rate is configured correctly (corresponding to expected return on investments).
6. Click on the "Run RRSP Optimization Button".
7. The model will run silently in a fraction of a second, automatically rewriting Column M to reveal your mathematically optimized deduction roadmap.
