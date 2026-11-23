WITH paired AS (
    SELECT a.economy, a.iso, a.value AS baseline, b.value AS comparison
    FROM observations AS a
    JOIN observations AS b ON b.iso = a.iso AND b.indicator = a.indicator AND b.year = 2023
    WHERE a.indicator = 'SP.URB.TOTL.IN.ZS' AND a.year = 2010
      AND a.value IS NOT NULL AND b.value IS NOT NULL
)
SELECT economy, iso, baseline, comparison, comparison - baseline AS metric
FROM paired ORDER BY metric DESC, iso;
