USE presupuesto;

CREATE OR REPLACE VIEW tabla_cruce AS
WITH eje_base AS (
    SELECT
        e.*,
        CONCAT_WS('-', e.tipo, e.cta, e.subc, e.objg, e.ord, e.sord, e.item) AS rubro_armado
    FROM eje e
),
cdp_agg AS (
    SELECT
        fecha_corte,
        rubro,
        fuente,
        recurso,
        sit,
        COUNT(*) AS cantidad_cdp,
        SUM(COALESCE(valor_actual, 0)) AS valor_cdp
    FROM cdp
    GROUP BY fecha_corte, rubro, fuente, recurso, sit
),
crp_agg AS (
    SELECT
        fecha_corte,
        rubro,
        fuente,
        recurso,
        situacion,
        COUNT(*) AS cantidad_crp,
        SUM(COALESCE(valor_actual, 0)) AS valor_crp
    FROM crp
    GROUP BY fecha_corte, rubro, fuente, recurso, situacion
)
SELECT
    e.id,
    e.fecha_corte,
    e.dependecia_de_afectacion_de_gastos,
    e.tipo,
    e.cta,
    e.subc,
    e.objg,
    e.ord,
    e.sord,
    e.item,
    e.sitem,
    e.rubro_armado AS rubro,
    e.concepto,
    e.fuente,
    e.situacion,
    e.rec,
    e.recurso,
    e.es_resumen,
    COALESCE(e.apropiacion_vigente_dep_gsto, 0) AS apropiacion_vigente,
    COALESCE(e.total_cdp_dep_gstos, 0) AS total_cdp_eje,
    COALESCE(e.apropiacion_disponible_dep_gsto, 0) AS apropiacion_disponible,
    COALESCE(e.total_compromiso_dep_gstos, 0) AS total_compromiso_eje,
    COALESCE(e.total_obligaciones_dep_gstos, 0) AS obligaciones_rubro_eje,
    COALESCE(c.cantidad_cdp, 0) AS cantidad_cdp,
    COALESCE(c.valor_cdp, 0) AS valor_cdp,
    COALESCE(r.cantidad_crp, 0) AS cantidad_crp,
    COALESCE(r.valor_crp, 0) AS valor_crp,
    CASE
        WHEN COALESCE(c.cantidad_cdp, 0) = 0 AND COALESCE(r.cantidad_crp, 0) = 0 THEN 0
        ELSE COALESCE(e.total_obligaciones_dep_gstos, 0)
    END AS obligaciones_para_semaforo,
    CASE
        WHEN COALESCE(c.cantidad_cdp, 0) = 0 AND COALESCE(r.cantidad_crp, 0) = 0 THEN 0
        WHEN COALESCE(e.apropiacion_vigente_dep_gsto, 0) = 0 THEN 0
        ELSE ROUND(COALESCE(e.total_obligaciones_dep_gstos, 0) / e.apropiacion_vigente_dep_gsto, 6)
    END AS semaforo
FROM eje_base e
LEFT JOIN cdp_agg c
    ON c.fecha_corte = e.fecha_corte
   AND c.rubro = e.rubro_armado
   AND c.fuente = e.fuente
   AND c.recurso = e.recurso
   AND c.sit = e.situacion
LEFT JOIN crp_agg r
    ON r.fecha_corte = e.fecha_corte
   AND r.rubro = e.rubro_armado
   AND r.fuente = e.fuente
   AND r.recurso = e.recurso
   AND r.situacion = e.situacion;
