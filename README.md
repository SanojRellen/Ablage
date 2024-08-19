
To illustrate how different types of autocallables (Standard, Memory, and Relax versions) are structured, I'll provide a hypothetical example for each. These examples will demonstrate the use of options to create the desired payoff structures. For simplicity, let's assume each product is linked to a single stock index.

### **1. Standard Autocallable Example**

#### **Structure:**
- **Underlying Asset:** Stock Index XYZ
- **Autocall Level:** 100% of initial index level
- **Barrier Level:** 70% of initial index level
- **Coupon:** 5% per quarter
- **Maturity:** 3 years
- **Observation Dates:** Quarterly

#### **Options Used:**
- **Call Options:** The issuer buys European call options with a strike price of 100% of the initial index level for each observation date. These calls ensure that if the index level is at or above 100% on an observation date, the note can be called early, paying the coupon and returning the principal.
- **Put Options (Barrier Options):** The issuer sells European put options with a strike price at the barrier level of 70% of the initial index level. These options define the barrier level and help finance the product. If the index falls below the barrier at maturity, the investor faces a loss proportional to the decline.

#### **Payoff Structure:**
- **Scenario 1 (Bull Market):** If the index is at or above 100% on any observation date, the note is called early, paying the quarterly coupon and returning the principal.
- **Scenario 2 (Sideways Market):** If the index never hits 100%, but stays above 70% at maturity, the investor receives the principal and accumulated coupons.
- **Scenario 3 (Bear Market):** If the index falls below 70% at maturity, the investor suffers a loss equivalent to the index's decline below the barrier.

### **2. Memory Express Autocallable Example**

#### **Structure:**
- **Underlying Asset:** Stock Index XYZ
- **Autocall Level:** 100% of initial index level
- **Barrier Level:** 70% of initial index level
- **Coupon:** 5% per quarter (with memory feature)
- **Maturity:** 3 years
- **Observation Dates:** Quarterly

#### **Options Used:**
- **Call Options:** The issuer buys European call options with a strike price of 100% of the initial index level for each observation date, just like in the standard version.
- **Put Options (Barrier Options):** The issuer sells European put options with a strike price at the barrier level of 70% of the initial index level, again defining the barrier.
- **Coupon Recovery (Memory) Feature:** Additional options or a mechanism is embedded to allow missed coupons to be paid retrospectively if subsequent conditions are met (e.g., a call option with a lower strike tied to the previous high-water mark of the index).

#### **Payoff Structure:**
- **Scenario 1 (Bull Market):** Similar to the standard version, if the index is at or above 100% on any observation date, the note is called early, and the investor receives the principal and the coupon. Additionally, if any previous coupons were missed, they are paid as well.
- **Scenario 2 (Sideways Market):** If the index doesn't reach 100% on the early observation dates but hits 100% later, the investor receives all the accumulated coupons that were previously missed, plus the principal.
- **Scenario 3 (Bear Market):** If the index falls below 70% at maturity, the investor suffers a loss equivalent to the index's decline below the barrier. Missed coupons may be forfeited if the index never recovers to trigger the memory feature.

### **3. Relax Autocallable Example**

#### **Structure:**
- **Underlying Asset:** Stock Index XYZ
- **Autocall Level:** 100% of initial index level
- **Barrier Level:** 60% of initial index level (lower than standard)
- **Coupon:** 4% per quarter (lower than standard due to more protection)
- **Maturity:** 3 years
- **Observation Dates:** Quarterly

#### **Options Used:**
- **Call Options:** The issuer buys European call options with a strike price of 100% of the initial index level for each observation date.
- **Put Options (Barrier Options):** The issuer sells European put options with a strike price at the barrier level of 60% of the initial index level. This lower barrier provides greater protection.
- **Airbag Feature:** Additional protective options (or structured notes) are embedded to limit losses if the index breaches the barrier but remains above a lower "airbag" level, say 50%. This feature could involve purchasing a deep out-of-the-money call option to offset some losses in extreme downside scenarios.

#### **Payoff Structure:**
- **Scenario 1 (Bull Market):** Similar to the standard version, the note is called early if the index is at or above 100% on an observation date, paying the coupon and returning the principal.
- **Scenario 2 (Sideways Market):** If the index stays above 60% (but below 100%), the investor receives the principal and accumulated coupons at maturity.
- **Scenario 3 (Bear Market):** If the index drops below 60% but stays above the airbag level of 50%, the investor's losses are cushioned, potentially receiving a partial return of principal or a reduced loss. If the index falls below the airbag level, the losses might be more significant but still somewhat protected compared to the standard version.

### **Summary**

- **Standard Autocallable:** Typically uses call options at 100% and put options at 70%. Offers straightforward autocall and barrier features without additional protective mechanisms.
- **Memory Express Autocallable:** Similar to standard but includes a memory feature, requiring additional options to enable the recovery of missed coupons if the index recovers.
- **Relax Autocallable:** Similar to the standard version but with a lower barrier and possibly an airbag feature, which requires different strike prices (e.g., 60% for the barrier and 50% for the airbag) and additional protective options to cushion downside risk.

Each variant is tailored to different market conditions and investor preferences, balancing the trade-off between protection, yield, and complexity.
