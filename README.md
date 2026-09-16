# fulfillment-productivity-analysis
Excel-based analysis of raw warehouse order fulfillment data to measure pick volume and productivity at the coworker level.

## Business Problem
Our raw order fulfillment reporting shows detailed picking data at the orderline-level, but does not provide an easy way to compare our fulfillment coworkers' picking performance. Because one single customer order can contain multiple orderlines, pick areas, coworkers, and tasks, simple row counts are not enough to produce insightful picking data analytics that enable us to track and manage coworker performance.

## Solution
I have built an **Excel-based analytics tool** using custom columns, formulas, and a pivot table to transform the raw data into metrics that show order picking statistics at the coworker level. 

## Dashboard and Analysis

<img width="1151" height="616" alt="image" src="https://github.com/user-attachments/assets/ba177083-55c7-4eab-9bf8-8efbdc89df22" />

*(Coworker IDs have been altered to protect personal and data privacy)*

The completed pivot table above brings the chosen metrics together to provide a coworker-level view of order fulfillment productivity. The report can be filtered by date, order type, and pick area, allowing performance to be analyzed across different periods and areas of the warehouse.

<img width="830" height="514" alt="image" src="https://github.com/user-attachments/assets/412d21d7-457e-4fa5-b3f6-15df6ed7387b" />

<img width="833" height="515" alt="image" src="https://github.com/user-attachments/assets/454a06a9-b16a-4c14-bdbc-08836ae54857" />

The charts provide two different views of coworker picking performance. Total weight picked provides context around overall workload, while orderlines per hour provides a standardized productivity metric that accounts for differences in time spent picking.

<img width="839" height="517" alt="image" src="https://github.com/user-attachments/assets/250b823a-ca54-494f-b720-75d41c7aca35" />

The scatter plot compares time spent picking with total orderlines picked, making it easier to view differences in productivity and outliers among coworkers that have similar amounts of picking activity.

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

## Identifying Unique Picking Tasks and Time Spent Picking
The raw data provides the time an order was started and the time each orderline was picked, but does not directly identify when an individual picking task was completed. Although the **Picked Status Time** shows when the entire order was completed, an order can be split across 3 areas of our warehouse and may be picked by multiple coworkers across these different areas. As a result, using this timestamp would provide inaccurate data on how long it took a coworker to complete a picking task.

To estimate the duration of each pick, I used **Time Started** as the starting point and the **Orderline Picked** timestamp of the final scanned orderline as the endpoint. And because a single order can generate multiple picking tasks, I incorporated the **Group Type** column to distinguish how our picking system divided the order into separate tasks. Considering these factors, a unique picking task is identified by using a combination of:

**User Picking (BF)** + **Order Number (D)** + **Pick Area (Q)** + **Group Type (I)**

```excel
=IF(AL2=MAXIFS($AL:$AL,
$BF:$BF,BF2,
$D:$D,D2,
$Q:$Q,Q2,
$I:$I,I2
),1,0)
```

The `MAXIFS` function of the formula locates the final **Orderline Picked (AL)** for each Unique Picking Task combination. The row containing this final scan returns a value of 1 in the **Unique Picking Task** column, identifying the endpoint of that picking task. Then, I created a **Time Spent Picking** column next to it and used the following formula to calculate the duration of each picking task:

```excel
=IF(AR2=1,AL2-AI2,0)
```

This subtracts the **Time Started (AI)** from the time on the last **Orderline Picked (AL),** giving us the total time from the moment the coworker went into the pick task until the last product on the pick task was scanned.

## Key Takeaways & Demonstrated Performance Management

We can see in the scatter plot that CW005 is an outlier when comparing **Time Spent Picking vs. Orderlines Picked**. While this coworker completed a similar number of orderlines as their peers, they recorded substantially more time spent picking. For our team, this has a negative impact on one of our major KPIs, which is **Orderlines per Hour**. 

After noticing this, I took a conversation with this coworker to see what was causing this difference. This coworker informed me that they often select several pick tasks at a time and then proceed to complete each task one by one. For example, if 10 picking tasks were selected at 9:00 AM, the first might be completed at 9:10 AM while the final task might not be completed until 11:00 AM. Because picking time is calculated from when each task is initially selected, the final task would appear in the data as having taken approximately two hours to complete.

I was able to use the data to explain to the coworker how this workflow was affecting their measured productivity and our Orderlines per Hour KPI. The coworker understood the impact and agreed to select and complete one designated picking task at a time going forward.

## Skills Demonstrated

- Intermediate Excel formulas and functions (`COUNTIFS`, `MAXIFS`, `IF`, `AND`)
- PivotTables and charts for data visualization and analysis
- Data cleaning and transformation
- Helper columns and conditional logic
- Productivity and operational performance analysis
- KPI development and performance measurement
- Identification and assessment of operational outliers
- Translating raw operational data into actionable insights
