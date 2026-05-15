# Container Calculator — Functional Design

---

## 1. Inputs

### 1 Product Master (Excel file)

The product master is the catalogue of all shippable products. The tool reads the first worksheet and requires the following columns (column names are matched flexibly, case-insensitively):

| Field | Required | Purpose |
|---|---|---|
| Product Code | Yes | Unique identifier (must be digits only) |
| Product Name | Yes | Human-readable label |
| Length (mm) | Yes | Case outer dimension |
| Width (mm) | Yes | Case outer dimension |
| Height (mm) | Yes | Case outer dimension |
| Gross Weight (kg) | Yes | Total weight of one case |
| Cases Per Pallet | Yes | Used by palletized mode |
| Max Stack | No | Maximum number of cases that can be stacked vertically |

**Validation:** Any row missing a product code, name, or any numeric dimension is silently skipped. If no valid rows remain after skipping, loading fails.



### 1.1 Loading Method

The user selects one of two modes before running the calculation:

| Mode | Description |
|---|---|
| **Loose Load** | Cases are stacked directly on the container floor. The tool plans exact floor positions and stack heights. |
| **Palletized** | Cases are grouped onto pallets. The tool allocates pallets to containers by count and weight, without planning floor positions. |

---

## 2. Container Types

The tool works with three standard dry container sizes:

| Type | Internal Length | Internal Width | Internal Height | CBM Limit | Weight Limit (kg) | Pallet Capacity (floor-loaded) |
|---|---|---|---|---|---|---|
| 20ft Standard | 5.90 m | 2.35 m | 2.39 m | 28 m³ | 21,770 | 10 |
| 40ft Standard | 12.03 m | 2.35 m | 2.39 m | 58 m³ | 26,730 | 20 |
| 40ft High Cube | 12.03 m | 2.35 m | 2.69 m | 68 m³ | 26,540 | 20 |

> **Note on CBM limits:** For the loose load mode, the CBM utilisation metric is calculated against these declared limits (not the raw geometric volume of the container). This reflects typical industry practice where unusable corner space is factored in.

---

## 3. Loose Load Algorithm

The loose load algorithm assigns the entire order to one or more containers, subject to physical space and weight limits. It proceeds container-by-container until all cases are assigned.

### 3.1 Container Selection Loop (`selectContainers`)

```
WHILE there are unassigned cases:
    TRY to fit all remaining cases into a single container, 
    testing sizes from smallest (20ft) to largest (40hc).

    IF a container size fits all remaining cases (space-wise):
        CHECK weight against that container's weight limit.
        IF weight is within limit:
            Assign all remaining cases to this container. Done.
        ELSE:
            Enforce the weight limit (see §4.5).
            Assign the within-limit portion to this container.
            The overflow becomes the new "remaining" for the next iteration.

    IF no single container size can fit all remaining cases:
        Greedily fill a 40hc container with as many cases as possible (see §4.4).
        The leftover cases become the new "remaining" for the next iteration.

    IF even a single case of the first remaining product cannot physically fit
    in any container:
        FAIL with an error — the product is physically incompatible.
```

**Key decision:** The algorithm always tries the smallest container first. This minimises cost when the entire order is small enough to fit in one box. It only escalates to larger containers when necessary.

### 4.3 Packing a Single Container (`packOrder`)

Given a fixed container and a list of items to pack, this function determines whether and how all the cases can be arranged physically.

#### Step 1 — Physical feasibility check

Before attempting any placement, the algorithm checks that every product's case height is ≤ the container's internal height. Any product that fails this check causes the packing to fail immediately with a descriptive error.

#### Step 2 — Build the stack list

Products are sorted by **quantity descending** (highest case count first). This heuristic places large-volume products first, giving them priority access to the best floor positions.

For each product, the algorithm calculates the **maximum stack height**:

```
maxByHeight  = floor(containerHeight / caseHeight)
effectiveMax = min(product.maxStackHeight, maxByHeight)   [if maxStackHeight is defined]
             = maxByHeight                                 [if maxStackHeight is not defined]
```

The product's cases are then divided into stacks of `effectiveMax` each (the last stack may be smaller if cases don't divide evenly). Each stack is treated as a single rectangular block for floor planning purposes.

#### Step 3 — Floor plan placement (Guillotine algorithm)

The container floor is modelled as a set of **free rectangles**. Initially there is one free rectangle covering the entire floor.

For each stack, the algorithm finds the **best-fit placement**:

1. For each free rectangle, test both orientations of the stack's footprint (length × width, and width × length).
2. Discard orientations that don't fit within the rectangle's bounds.
3. Among valid placements, select the one that:
   - Is **furthest from the door** (highest x-coordinate first), then
   - Furthest to one side (highest y-coordinate), then
   - Has the **least wasted area** (tightest fit).

> **Business rationale for loading from the back:** Filling the deepest positions first reflects real-world container loading practice — goods loaded first end up furthest from the doors.

4. Place the stack at the chosen position.
5. **Split the used rectangle** into two remaining free rectangles:
   - A rectangle to the **right** of the placed stack (same depth, remaining width).
   - A rectangle **above** the placed stack (full depth of the original rect, remaining height).

If no free rectangle can accommodate a stack, packing fails — the order exceeds this container's floor area.

#### Output metrics

| Metric | Calculation |
|---|---|
| Total cases | Sum of all case counts in the order |
| Total weight (kg) | Sum of (gross weight per case × cases in stack) for all stacks |
| Volume used (m³) | Sum of (footprint area × stack height) for all stacks |
| CBM utilisation (%) | Volume used ÷ container CBM limit × 100 |
| Stack count | Number of stacks placed |

### 4.4 Greedy Fill (`greedyFill`)

Used when the total order does not fit in any single container. The goal is to fill a 40hc container with as many cases as possible, respecting both space and weight.

```
FOR each product (in code order):
    TRY adding the full quantity of this product to the packed set.
    IF it fits (space and weight):
        Add it.
    ELSE:
        BINARY SEARCH for the maximum number of cases of this product that fit.
        Add that many cases to this container.
        Put the remainder in overflow.
```

The binary search runs `packOrder` at each step to test feasibility. This is computationally intensive for large orders but ensures an accurate answer.

### 4.5 Weight Limit Enforcement (`enforceWeightLimit`)

Used when an entire order's cases fit spatially in a container but the combined weight exceeds the limit. The algorithm removes cases from the last product first:

```
WHILE total weight > limit:
    TRY binary searching for the maximum cases of the last product that fit within the weight limit.
    IF a partial quantity fits:
        Keep that many in this container; move the rest to overflow.
        STOP.
    ELSE:
        Remove the entire last product from this container; move it to overflow.
        Repeat.
```

---

## 5. Palletized Load Algorithm

### 5.1 Pallet Generation

Before assigning pallets to containers, the algorithm converts the order into a list of individual pallets:

```
FOR each product in the order:
    DIVIDE the case quantity by cases_per_pallet.
    Each full group of cases_per_pallet becomes one full pallet.
    Any remainder becomes a partial pallet.
```

Each pallet carries its weight as: `(cases on pallet × gross weight per case) + 20 kg` (the empty pallet tare weight).

> **Note on partial pallets:** A partial pallet is treated exactly like a full pallet for allocation purposes — it still occupies one pallet slot in the container.

### 5.2 Container Assignment Loop (`selectPalletContainers`)

Pallets are sorted by product code before assignment, keeping the same product together where possible.

```
WHILE there are unassigned pallets:
    TRY fitting all remaining pallets into the smallest container possible,
    testing 20ft → 40ft → 40hc in order.
    A container "fits" if:
        remaining pallet count ≤ container's pallet capacity  AND
        total pallet weight ≤ container's weight limit.

    IF a container fits all remaining pallets:
        Assign all to that container. Done.

    IF no container fits all remaining pallets:
        Use a 40hc container.
        Load pallets one by one until either the capacity (20 pallets) or
        the weight limit is reached.
        Assign loaded pallets to this container.
        The rest become "remaining" for the next iteration.

    IF even a single pallet exceeds the weight limit of a 40hc:
        FAIL with an error.
```

**Key difference from loose load:** Palletized mode does not perform floor-level space optimisation. It treats containers as having a fixed pallet slot count (a GMA pallet standard) and checks only count and weight. There is no 2D or 3D visualisation for palletized loads.

---

## 6. Results and Output

### 6.1 Summary Card

Always displayed, regardless of loading mode. Shows:
- Total number of containers required
- Total cases packed
- Total pallets (palletized mode only)
- Total gross weight across all containers
- A per-container breakdown table

### 6.2 Volume Utilisation (Loose Mode Only)

CBM utilisation is colour-coded in the results:
- **Below 80%** — displayed in red (poor utilisation; consider consolidation)
- **80% and above** — displayed in brand blue (acceptable)

### 6.3 Per-Container Detail Cards (Loose Mode Only)

For each container:
- **Floor plan** — a 2D top-down SVG diagram showing each stack's footprint, colour-coded by product. Doors are shown at the front (right) end.
- **3D view** — an isometric SVG projection showing stacking heights and door positions.
- **Placement detail table** — expandable rows, one per product, showing individual stack dimensions, case count, height, and weight.

For palletized loads, only the summary card is shown. The individual container cards are suppressed because pallet-level detail is fully represented in the summary table.

---

## 7. Validation and Error Handling

| Condition | Result |
|---|---|
| Product master missing required columns | Upload fails with column list |
| No valid rows in product master | Upload fails |
| SKU contains non-digit characters | Order line rejected with warning |
| SKU not found in catalogue | Order line rejected with warning |
| Order is empty when Containerize is clicked | Error message |
| Order contains SKUs not in catalogue | Error message; calculation blocked |
| Palletized mode with missing cases-per-pallet data | Error message; calculation blocked |
| Case height exceeds container height | Packing fails with product name |
| Product cannot fit in any container physically | Allocation fails with product name |
| Single pallet exceeds weight limit | Palletized allocation fails with product name |

---

## 8. Key Business Rules Summary

1. **Container size preference:** Always use the smallest container that fits the remaining load. Escalate to larger containers only when necessary.
2. **Loading direction:** In loose mode, stacks are placed from the back of the container toward the doors, matching real-world loading practice.
3. **Product grouping:** Products are kept together within a container where possible (same-product stacks are generated consecutively). This simplifies physical unloading.
4. **Stacking limits:** If a product master specifies a maximum stack height, it is honoured and takes precedence over the height-derived limit. This reflects fragility or stacking strength constraints.
5. **Weight always wins:** A container that passes the space check can still be split if it fails the weight check. Weight limits are hard constraints.
6. **Pallets have tare weight:** The 20 kg empty pallet weight is included in all palletized weight calculations, ensuring the weight limit reflects real-world shipping weights.
7. **Partial pallets count as full slots:** A partial pallet occupies the same container slot count as a full pallet — consistent with how carriers charge for pallet positions.
