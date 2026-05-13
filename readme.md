# Container Optimization Algorithm Logic

This document describes the container packing algorithm implemented in the application. The algorithm is responsible for selecting the appropriate containers and packing products into them while respecting physical constraints (dimensions and weight limits).

## Overview
The optimization process consists of two main levels:
1. **Container Selection & Allocation (`selectContainers`)**: Determines the mix of containers (20ft, 40ft, 40hc) and partitions the order among them.
2. **Container Packing (`packOrder`)**: Determines if a specific set of items can physically fit into a single container's dimensions using a 2D bin-packing approach with height constraints.

## 1. Container Selection and Allocation (`selectContainers`)

```text
FUNCTION selectContainers(orderItems):
    remainingItems = copy of orderItems
    sort remainingItems by product code alphabetically
    containers = empty list
    containerTypes = ['20ft', '40ft', '40hc']

    WHILE remainingItems is not empty:
        selectedType = NULL
        packResult = NULL

        // Try to fit all remaining items into a single container, starting from the smallest
        FOR each type IN containerTypes:
            result = packOrder(remainingItems, type.dimensions)
            IF result is successful:
                selectedType = type
                packResult = result
                BREAK

        IF selectedType is found:
            // All remaining items physically fit into selectedType
            IF packResult.total_weight > type.weightLimit:
                // Weight limit exceeded, remove items until it's under the limit
                packed, overflow = enforceWeightLimit(remainingItems, selectedType)
                result = packOrder(packed, selectedType.dimensions)
                add {type: selectedType, result: result} to containers
                remainingItems = overflow
            ELSE:
                // Fits both physically and by weight
                add {type: selectedType, result: packResult} to containers
                remainingItems = empty
        ELSE:
            // Items do not fit in any single container, greedily fill the largest ('40hc')
            packed, overflow = greedyFill(remainingItems, '40hc')
            
            IF packed is empty:
                // Edge case: A single item cannot fit
                RETURN error "Product cannot physically fit"
            
            result = packOrder(packed, '40hc'.dimensions)
            add {type: '40hc', result: result} to containers
            remainingItems = overflow

    RETURN success, containers
```

## 2. Greedy Fill (`greedyFill`)
Fills a specific container type as much as possible with the remaining items sequentially.

```text
FUNCTION greedyFill(items, type):
    packed = empty list
    overflow = empty list

    FOR each item IN items:
        trialPack = packed + item
        result = packOrder(trialPack, type.dimensions)
        
        IF result is successful AND result.total_weight <= type.weightLimit:
            add item to packed
        ELSE:
            // Binary search to find the maximum number of cases of this item that will fit
            low = 0
            high = item.case_count
            bestCases = 0
            
            WHILE low <= high:
                mid = floor((low + high) / 2)
                IF mid == 0:
                    low = 1
                    CONTINUE
                
                trialPack = packed + (item with mid cases)
                result = packOrder(trialPack, type.dimensions)
                
                IF result is successful AND result.total_weight <= type.weightLimit:
                    bestCases = mid
                    low = mid + 1
                ELSE:
                    high = mid - 1
                    
            IF bestCases > 0:
                add (item with bestCases) to packed
                add (item with remaining cases) to overflow
            ELSE:
                add item to overflow

    RETURN packed, overflow
```

## 3. Enforce Weight Limit (`enforceWeightLimit`)
Removes items from the end of the list until the total weight is under the container's weight limit.

```text
FUNCTION enforceWeightLimit(items, type):
    packed = copy of items
    overflow = empty list

    WHILE packed is not empty:
        result = packOrder(packed, type.dimensions)
        IF result is successful AND result.total_weight <= type.weightLimit:
            RETURN packed, overflow
            
        lastItem = last element in packed
        IF lastItem.case_count > 1:
            // Binary search to find maximum cases that keep weight under limit
            low = 1
            high = lastItem.case_count
            bestCases = 0
            
            WHILE low <= high:
                mid = floor((low + high) / 2)
                trialPack = packed (without lastItem) + (lastItem with mid cases)
                result = packOrder(trialPack, type.dimensions)
                
                IF result is successful AND result.total_weight <= type.weightLimit:
                    bestCases = mid
                    low = mid + 1
                ELSE:
                    high = mid - 1
            
            IF bestCases > 0:
                add (lastItem with remaining cases) to start of overflow
                update lastItem in packed to have bestCases
                RETURN packed, overflow
                
        // If we reach here, either it was 1 case or even 1 case was too heavy
        remove lastItem from packed
        add lastItem to start of overflow

    RETURN packed, overflow
```

## 4. Container Packing (`packOrder`)
Determines if a list of items can fit into a specific container size using a 2D Best-Area-Fit heuristic with 1D height stacking.

```text
FUNCTION packOrder(order, container):
    // 1. Initial height check
    IF any item.case_height > container.height:
        RETURN failure "Item too tall"

    // 2. Sort items by base area (length * width) descending
    sortedOrder = sort order by (case_length * case_width) descending

    // 3. Form stacks
    stackList = empty list
    FOR each item IN sortedOrder:
        maxByHeight = floor(container.height / item.case_height)
        effectiveMax = maxByHeight
        IF item.max_stack_height is set:
            effectiveMax = minimum(item.max_stack_height, maxByHeight)
        
        remainingCases = item.case_count
        WHILE remainingCases > 0:
            casesInThisStack = minimum(remainingCases, effectiveMax)
            create stack = {
                code: item.code,
                length: item.case_length,
                width: item.case_width, 
                cases_in_stack: casesInThisStack,
                physical_height: casesInThisStack * item.case_height,
                weight: casesInThisStack * item.gross_weight
            }
            add stack to stackList
            remainingCases -= casesInThisStack

    // 4. Floor placement (Guillotine Split Heuristic)
    floorPlan = empty list
    freeRects = [{x: 0, y: 0, length: container.length, width: container.width}]

    FOR each stack IN stackList:
        bestRect = NULL
        bestWaste = Infinity
        bestOrientation = NULL

        // Find the best free rectangle to place this stack
        FOR each rect IN freeRects:
            FOR each orientation IN [(length, width), (width, length)]:
                IF stack dimensions (in orientation) <= rect dimensions:
                    waste = (rect.length * rect.width) - (stack.length * stack.width)
                    IF waste < bestWaste:
                        bestWaste = waste
                        bestRect = rect
                        bestOrientation = orientation

        IF bestRect is NULL:
            RETURN failure "Insufficient floor space"

        // Place stack and split remaining space
        place stack at (bestRect.x, bestRect.y)
        add to floorPlan
        
        // Remove the chosen free rect
        remove bestRect from freeRects
        
        // Split the remaining space into two new free rectangles (Right and Top)
        rightRect = {
            x: bestRect.x + stack.placedLength, 
            y: bestRect.y, 
            length: bestRect.length - stack.placedLength, 
            width: stack.placedWidth
        }
        topRect = {
            x: bestRect.x, 
            y: bestRect.y + stack.placedWidth, 
            length: bestRect.length, 
            width: bestRect.width - stack.placedWidth
        }
        
        IF rightRect.area > 0: add rightRect to freeRects
        IF topRect.area > 0: add topRect to freeRects

    RETURN success, floorPlan, placementMetrics
```
