# Excel Bond/Security Financial Functions — Mathematical Formulas for TypeScript Implementation

Reference formulas for implementing Excel-compatible bond and security functions. All parameters use Excel conventions.

---

## 1. Day Count Conventions (basis parameter 0–4)

**Basis mapping:**

| Basis | Convention | Year (B) | Description |
|-------|------------|----------|-------------|
| 0 | US (NASD) 30/360 | 360 | Month = 30 days, special rules for month-end |
| 1 | Actual/Actual | 365 or 366 | Actual days; split at year boundary for leap years |
| 2 | Actual/360 | 360 | Actual days / 360 |
| 3 | Actual/365 | 365 | Actual days / 365 |
| 4 | European 30/360 | 360 | Like 30/360 but D2=31 → 30 always |

### 30/360 base formula (both US and European)

```
days = (Y2 - Y1) × 360 + (M2 - M1) × 30 + (D2 - D1)
```

Where `(D1, M1, Y1)` = start date, `(D2, M2, Y2)` = end date.

### US 30/360 (basis 0) — adjustments before applying formula

- If D1 = 31 → set D1 = 30
- If D2 = 31 and D1 ∈ {30, 31} → set D2 = 30
- If D1 = last day of February → set D1 = 30
- If D2 = last day of February and D1 was last day of February → set D2 = 30

### European 30/360 (basis 4) — adjustments

- If D1 = 31 → set D1 = 30
- If D2 = 31 → set D2 = 30

### Actual/Actual (basis 1) — ISDA-style

Split the period at year boundaries. For each sub-period:

- Days in leap year → divide by 366
- Days in non-leap year → divide by 365

```
fraction = Σ (days_in_subperiod_i / (365 or 366 for that year))
```

### Actual/360 (basis 2)

```
fraction = actual_days / 360
```

### Actual/365 (basis 3)

```
fraction = actual_days / 365
```

### Helper: YEARFRAC(start, end, basis)

Returns the year fraction between two dates using the given basis. Use the same day-count logic as above.

---

## 2. Coupon Date Functions

**Shared parameters:** `settlement`, `maturity`, `frequency` (1/2/4), `basis` (0–4)

**Coupon period:** Interval between coupon dates. For frequency `f`:
- `periodMonths = 12 / f` (12 for annual, 6 for semiannual, 3 for quarterly)

**Coupon dates:** Align with maturity. Maturity is always a coupon date. Other dates are `maturity - k × periodMonths` for k = 1, 2, 3, …

### Algorithm: COUPPCD (previous coupon date)

Largest coupon date ≤ settlement.

1. Start from `maturity`.
2. While `maturity - k × periodMonths > settlement`, increment k.
3. `COUPPCD = maturity - k × periodMonths` (date arithmetic in months).

### Algorithm: COUPNCD (next coupon date)

Smallest coupon date > settlement.

1. `COUPPCD = COUPPCD(settlement, maturity, frequency, basis)`
2. `COUPNCD = COUPPCD + periodMonths`

### COUPDAYBS (days from coupon start to settlement)

```
COUPDAYBS = days(COUPPCD, settlement, basis)
```

### COUPDAYS (days in coupon period containing settlement)

```
COUPDAYS = days(COUPPCD, COUPNCD, basis)
```

### COUPDAYSNC (days from settlement to next coupon)

```
COUPDAYSNC = days(settlement, COUPNCD, basis)
```

### COUPNUM (number of coupons from settlement to maturity)

Count coupon dates in `(settlement, maturity]` (exclude settlement, include maturity).

1. `ncd = COUPNCD(settlement, maturity, frequency, basis)`
2. Count coupons from `ncd` to `maturity` at the given frequency.
3. `COUPNUM = 1 + (maturity - ncd) / periodMonths` (in periods).

---

## 3. ACCRINT — Accrued Interest

```
ACCRINT(issue, first_interest, settlement, rate, par, frequency, [basis], [calc_method])
```

**Parameters:**
- `issue` — issue date
- `first_interest` — first coupon date
- `settlement` — settlement date
- `rate` — annual coupon rate (e.g. 0.05 for 5%)
- `par` — par value (default 1000)
- `frequency` — 1, 2, or 4
- `basis` — 0–4 (default 0)
- `calc_method` — TRUE: issue→settlement; FALSE: first_interest→settlement (default TRUE)

**Formula (calc_method = TRUE):**

Accrued interest from issue to settlement using quasi-coupon periods in any odd first period:

```
ACCRINT = par × rate × Σ(Aj / NLj) / NC
```

Where:
- `NC` = number of quasi-periods in the odd period (issue to first_interest)
- `Aj` = accrued days in quasi-period j
- `NLj` = length of quasi-period j (in days, per basis)

**Simpler approach (typical case):**

1. `COUPPCD =` previous coupon before/on settlement (use maturity = first_interest or a proxy).
2. For standard bonds, treat `first_interest` as anchor and build schedule.
3. `A = days(COUPPCD, settlement, basis)`
4. `E = days(COUPPCD, COUPNCD, basis)`
5. `coupon_per_period = par × rate / frequency`
6. `ACCRINT = coupon_per_period × (A / E)`

For the first (odd) period, use quasi-coupon periods from issue to first_interest.

---

## 4. PRICE — Bond Price

```
PRICE(settlement, maturity, rate, yld, redemption, frequency, [basis])
```

**Parameters:**
- `rate` — annual coupon rate
- `yld` — annual yield to maturity
- `redemption` — redemption per 100 face (often 100)
- `frequency` — 1, 2, or 4

**Auxiliary values (using basis for day count):**
- `DSC` = days(settlement, COUPNCD, basis)
- `E` = days(COUPPCD, COUPNCD, basis)
- `A` = days(COUPPCD, settlement, basis)
- `N` = number of coupons from settlement to maturity (COUPNUM)

**When N > 1:**

```
PRICE = [redemption / (1 + yld/freq)^(N - 1 + DSC/E)]
      + Σ(k=1 to N) [100 × rate/freq / (1 + yld/freq)^(k - 1 + DSC/E)]
      - 100 × rate/freq × A/E
```

The sum is the PV of coupons. The last term subtracts accrued interest (clean price).

**When N = 1:**

```
PRICE = (100 × rate/freq + redemption) / (1 + yld/freq)^(DSC/E) - 100 × rate/freq × A/E
```

(Compound discounting for the fractional period; some implementations use simple: `1 + yld/freq × DSC/E` — Excel uses compound.)

**Compact form (N > 1):**

Let `y = yld/freq`, `c = 100 × rate/freq`:

```
PRICE = redemption × (1+y)^(-(N-1+DSC/E))
      + c × [1 - (1+y)^(-N)] / [y × (1+y)^(DSC/E)]
      - c × A/E
```

---

## 5. YIELD — Bond Yield from Price

```
YIELD(settlement, maturity, rate, pr, redemption, frequency, [basis])
```

**When N ≤ 1:** Solve the single-coupon PRICE equation for `yld`.

**When N > 1:** Newton-Raphson on the PRICE formula.

```
y_new = y_old - (PRICE(y_old) - pr) / (dPRICE/dy)
```

- `PRICE(y)` = bond price at yield `y`
- `dPRICE/dy` = derivative of PRICE w.r.t. yield

**Derivative of PRICE w.r.t. y (per-period yield):**

For `P = Σ CF_k × (1+y)^(-t_k) - accrued`:

```
dP/dy = -Σ CF_k × t_k × (1+y)^(-t_k-1)
```

Use `y = yld/freq` and scale the derivative by `1/freq` for annual yield.

**Algorithm:**
1. `y0 = rate` (initial guess)
2. For i = 0..99: `y_{i+1} = y_i - (PRICE(y_i) - pr) / dPRICE_dy(y_i)`
3. Stop when `|PRICE(y_i) - pr| < ε` (e.g. 1e-7)
4. Return `y × frequency` as annual yield

---

## 6. DURATION / MDURATION

### Macaulay Duration

```
DURATION = Σ (t × PV(CF_t)) / Price
```

Where:
- `t` = time to cash flow (in years or periods)
- `PV(CF_t)` = present value of cash flow at time t
- `Price` = bond price (dirty price)

**Formula:**
```
DURATION = [Σ(k=1 to N) (t_k × C × (1+y)^(-t_k)) + t_N × FV × (1+y)^(-t_N)] / P
```

- `C` = coupon per period
- `FV` = redemption
- `y` = yield per period
- `t_k` = time to k-th coupon in years: `t_k = (k - 1 + DSC/E) / frequency`
- `P` = dirty price (clean price + accrued interest)

### Modified Duration

```
MDURATION = DURATION / (1 + yld/frequency)
```

`yld` = annual yield. Modified duration approximates % price change per 1% yield change.

---

## 7. DISC / INTRATE / RECEIVED (Discount Securities)

### DISC — Discount Rate

```
DISC(settlement, maturity, pr, redemption, [basis])
```

```
discount = (redemption - pr) / redemption × B / DSM
```

- `B` = days per year from basis
- `DSM` = days(settlement, maturity, basis)

### INTRATE — Interest Rate

```
INTRATE(settlement, maturity, investment, redemption, [basis])
```

```
rate = (redemption - investment) / investment × B / DIM
```

- `DIM` = days(settlement, maturity, basis)

### RECEIVED — Amount at Maturity

```
RECEIVED(settlement, maturity, investment, discount, [basis])
```

```
received = investment / (1 - discount × DIM / B)
```

- `investment` = amount paid at settlement (price)
- `DIM` = days(settlement, maturity, basis)

---

## 8. Treasury Bill Functions

T-Bills use a 360-day year. `DSM` = days(settlement, maturity) on a 360-day basis.

### TBILLPRICE

```
TBILLPRICE(settlement, maturity, discount)
```

```
price = 100 × (1 - discount × DSM / 360)
```

### TBILLYIELD

```
TBILLYIELD(settlement, maturity, pr)
```

```
yield = (100 - pr) / pr × 360 / DSM
```

### TBILLEQ — Bond-Equivalent Yield

```
TBILLEQ(settlement, maturity, discount)
```

```
TBILLEQ = (365 × discount) / (360 - discount × DSM)
```

Or equivalently, first get price, then:

```
TBILLEQ = (365 × yield_360) / (360 - yield_360 × DSM)
```

where `yield_360 = (100 - price) / price × 360 / DSM`.

---

## TypeScript Implementation Notes

1. **Dates:** Use a consistent date type (e.g. `Date` or serial days since epoch). Excel uses serial numbers; ensure your `days()` function matches the chosen basis.
2. **Month arithmetic:** For COUPPCD/COUPNCD, add/subtract months carefully (e.g. Jan 31 + 1 month → Feb 28/29).
3. **PRICE/YIELD:** PRICE returns clean price. For dirty price, add ACCRINT.
4. **Newton-Raphson:** Use a small tolerance (e.g. 1e-9) and cap iterations (e.g. 100).
5. **Basis 1 (Actual/Actual):** Implement ISDA-style splitting at year boundaries for leap years.
