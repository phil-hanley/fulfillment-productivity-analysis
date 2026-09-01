# fulfillment-productivity-analysis
Excel-based analysis of raw warehouse order fulfillment data to measure pick volume and productivity at the coworker level.

## Business Problem
Our raw order fulfillment reporting shows detailed picking data at the orderline-level, but does not provide an easy way to compare our fulfillment coworkers' picking performance. Because one single customer order can contain multiple orderlines, pick areas, coworkers, tasks, and simple row counts are not enough to produce insightful picking data analytics that enable us to track and manage coworker performance.

## Solution
I have built an **Excel based analytics tool** using custom columns, formulas, and a pivot table to transform the raw data into metrics that show order picking statistics at the coworker level. 

## Identifying Unique Orders Picked
As stated, the raw order fulfillment reporting is detailed at the orderline level, meaning a single order can contain several rows representing different products. If a coworker picks five products for one order, this should be represented as one order, not five.
To account for this, I created a **Unique Orders Picked** column in the raw data using `COUNTIFS` to analyze the **order number** and **user picking** columns. The first occurrence of each coworker-order combination is assigned a value of 1, while all subsequent occurrences are assigned 0. This column, which is then summed in the values field of the pivot table, provides the number of unique orders each coworker contributed to.

```excel
=IF(COUNTIFS($BF$2:BF2,BF2,$D$2:D2,D2)=1,1,0)
```
The expanding ranges of column **BF (user picking)** and column **D (order number)** allows Excel to determine whether the current coworker-order combination has already appeared previously in the dataset.

## Identifying Unique Picking Days
In order to measure how many days each coworker performed a pick, I created a helper column that identifies the first occurrence of each unique coworker-date combination.
During validation, I found that orders with a `Returned` status could appear with activity dates later than the original picking date. While these orders do represent a pick that the coworker completed, counting the return activity as a picking day inflated the metric. To account for this, I modified the calculation to exclude records with an order status of `Returned` **only when identifying unique picking days.** This allows for a proper count of unique picking days while preserving other metrics like orders, orderlines, and weight.

```excel
=IF(P2="Returned",0,IF(COUNTIFS($BF$2:BF2,BF2,$AJ$2:AJ2,AJ2,$P$2:P2,"<>Returned")=1,1,0))
```
This formula functions similarly to the Unique Orders Picked formula, but includes additional conditions to account for returned orders. Column **P (actual order status)** is evaluated to return 0 if the status is `Returned`, while the expanding ranges of column **BF (user picking)** and column **AJ (date orderline picked)** identify the first occurrence of each coworker-date combination. Returned records are ignored by the unique day logic to ensure that legitimate picking days are still counted, regardless of whether a returned order is associated with the same coworker-date combination.
