WITH RankedLotes AS (
    SELECT
        UPPER(tdus.no_unidade_saude) AS "UNIDADE",
        COALESCE(trl.nu_lote_originadora, 0) AS "LOTE",
        trl.dt_recebimento AS "DATE",
        trl.dt_recebimento AS "HORA DE IMPORTACAO",
        trl.no_responsavel_envio AS "ENVIADO POR",
        COALESCE(COUNT(CASE WHEN trve.co_seq_receb_valid_erro IS NULL THEN 1 END), 0) AS "FICHAS VALIDAS",
        COALESCE(COUNT(CASE WHEN trve.co_seq_receb_valid_erro IS NOT NULL THEN 1 END), 0) AS "FICHAS INVALIDOS",
        ttodt.no_tipo_origem_dado_transp AS "TIPO DE IMPORTACAO",
        ROW_NUMBER() OVER (PARTITION BY COALESCE(trl.nu_lote_originadora, 0) ORDER BY trl.dt_recebimento DESC) AS RowNum
    FROM
        tb_recebimento_lote trl
    INNER JOIN
        tb_tipo_origem_dado_transp ttodt ON trl.tp_origem_dado_recebido = ttodt.co_seq_tipo_origem_dado_transp
    INNER JOIN
        tb_recebimento_item tri ON trl.co_seq_receb_lote = tri.co_receb_lote
    INNER JOIN
        tb_dim_unidade_saude tdus ON tdus.nu_cnes = tri.nu_cnes_dado_serializado
    LEFT JOIN
        tb_recebimento_validacao_erros trve ON trve.co_receb_item = tri.co_seq_receb_item
    WHERE
        trl.dt_recebimento >= CURRENT_DATE - INTERVAL '6 months'
    GROUP BY
        tdus.no_unidade_saude,
        COALESCE(trl.nu_lote_originadora, 0),
        trl.dt_recebimento,
        trl.no_responsavel_envio,
        ttodt.no_tipo_origem_dado_transp
)
SELECT
    "UNIDADE",
    "LOTE",
    "DATE" AT TIME ZONE 'America/Fortaleza' AS "DATA",
    TO_CHAR("HORA DE IMPORTACAO" AT TIME ZONE 'America/Fortaleza', 'HH24:MI') AS "HORA DE IMPORTACAO",
    TO_CHAR("DATE" AT TIME ZONE 'America/Fortaleza', 'MM/YYYY') AS "PERIODO",
    "ENVIADO POR",
    "FICHAS VALIDAS",
    "FICHAS INVALIDOS",
    "TIPO DE IMPORTACAO"
FROM
    RankedLotes
WHERE
    RowNum = 1
ORDER BY
    "DATE" DESC;
