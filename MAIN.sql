SELECT
  ORDER,
  STATUS,
  REPLACE(LISTAGG(SKU, ', '), '   ', '') AS SKU,
  GLN,
  LISTAGG(CAST(SKU_QTY AS INT), ', ') AS SKU_QTY,
  OT_FLAG,
  SVIA,
  TOTAL_QTY,
  WEIGHT,
  TIMESTAMP(
    date(
      substr(char(ESD), 1, 4) || '-' || substr(char(ESD), 5, 2) || '-' || substr(char(ESD), 7, 2)
    )
  ) AS ESD,
  SCODE,
  DL_DATE_CST,
  COUNTRY,
  ORDER_TYPE,
  UPC_FLAG,
  DISTI_FLAG,
  ORDER_HOLD,
  CFI,
  (
    CASE WHEN LEFT(GIT, 3) = 'GIT' THEN 'GIT' WHEN GLN = 'DG' THEN 'DG' WHEN CFI = 'CF' THEN 'CFI' WHEN DISTI_FLAG = 'DS'
    AND UPC_FLAG = 'UPC' THEN 'DISTI - UPC' WHEN UPC_FLAG = 'UPC' THEN 'NON-DISTI - UPC' WHEN PLT_FLAG = 'Y' THEN 'PALLETIZER' WHEN ORDER_TYPE IN ('HD', 'HDB') THEN 'HD' WHEN COUNTRY = 'CA'
    AND DISTI_FLAG = 'DS' THEN 'DISTI - CA' WHEN COUNTRY = 'INTL'
    AND DISTI_FLAG = 'DS' THEN 'DISTI - INTL' WHEN COUNTRY = 'DOM'
    AND (
      DISTI_FLAG = 'DS'
      OR SCODE = 'A'
    ) THEN 'DISTI - DOM' WHEN COUNTRY <> 'DOM' THEN COUNTRY WHEN SCODE = 'FG1' THEN 'FG1' WHEN SVIA IN (
      'AACT',
      'AIIH',
      'AIT2',
      'AITL',
      'AITM',
      'AITO',
      'AVRT',
      'CEV1',
      'CEV2',
      'CEV3',
      'CEVG',
      'CV1O',
      'CV2O',
      'CV3O',
      'CVGO',
      'DZO1',
      'DZO2',
      'DZO3',
      'EXDO',
      'FAX1',
      'FAX2',
      'FAX3',
      'FDXO',
      'FEXF',
      'FH1O',
      'FH2O',
      'FH3O',
      'FXA1',
      'FXA2',
      'FXA3',
      'FXF1',
      'FXFE',
      'FXFO',
      'FXLO',
      'FXNL',
      'HWA',
      'MA1O',
      'MA2O',
      'MA3O',
      'MAC1',
      'MAC2',
      'MAC3',
      'MACV',
      'MCVO',
      'MN3O',
      'MOAO',
      'MOAV',
      'MTLE',
      'MTLO',
      'MV3O',
      'ODFL',
      'ODFO',
      'ODFR',
      'OFRO',
      'OVEO',
      'PAAM',
      'PAFO',
      'PAXP',
      'PGAA',
      'PILT',
      'PRLT',
      'RBTO',
      'RBTW',
      'RDFO',
      'RDWO',
      'RLCR',
      'SAIA',
      'SAIO',
      'SLCL',
      'SLCO',
      'SLCY',
      'UFO1',
      'UPSF',
      'USPS',
      'VALO',
      'VALS',
      'VC3O',
      'VS1O',
      'VS2O',
      'VS3O',
      'VSCO',
      'VSQO',
      'VST1',
      'VST2',
      'VST3',
      'VSTL',
      'VSTQ',
      'XPOL',
      'XPOO'
    )
    OR WEIGHT >= '419' THEN 'BULK' WHEN ORDER_TYPE IN ('TON', 'LZR') THEN 'TONER' WHEN ORDER_TYPE IN ('SNG', 'OEM', 'INK', 'PWC') THEN ORDER_TYPE ELSE 'PNP' END
  ) AS LOB,
  HOLD,
  BD_FLAG
FROM(
    SELECT
      PHPICK00.PHPKTN AS ORDER,
      PHPICK00.PHPSTF AS STATUS,
      PDPICK00.PDSTYL AS SKU,
      PHPICK00.PHGLN AS GLN,
      PDPICK00.PDPIQT AS SKU_QTY,
      PHPICK00.PHOTYP AS OT_FLAG,
      PHPICK00.PHSVIA AS SVIA,
      PHPICK00.PHTUTS AS TOTAL_QTY,
      ROUND(PHPICK00.PHESWT) AS WEIGHT,
      (
        CASE WHEN PHPICK00.PHDIDT <> '0' THEN PHPICK00.PHDIDT ELSE PHPICK00.PHCMDT END
      ) AS ESD,
      PHPICK00.PHMIS1 AS SCODE,
      TIMESTAMP(
        date(
          substr(char(PHPICK00.PHPDCR), 1, 4) || '-' || substr(char(PHPICK00.PHPDCR), 5, 2) || '-' || substr(char(PHPICK00.PHPDCR), 7, 2)
        ),
        TIME(
          SUBSTR(RIGHT('00' || PHPICK00.PHPTCR, 6), 1, 2) || '.' || SUBSTR(RIGHT('00' || PHPICK00.PHPTCR, 4), 1, 2) || '.' || RIGHT('00' || PHPICK00.PHPTCR, 2)
        )
      ) - 1 hours AS DL_DATE_CST,
      (
        CASE WHEN PHPICK00.PHSHCN = '124' THEN 'CA' WHEN PHPICK00.PHMIS1 IN ('F', ':', '3') THEN 'INTL' WHEN PHPICK00.PHSHCN = '840' THEN 'DOM' ELSE 'INTL' END
      ) AS COUNTRY,
      PHPICK00.PHMS12 AS GIT,
      PHPICK00.PHPKTS AS ORDER_TYPE,
      PHPICK00.PHAPCR AS UPC_FLAG,
      PHPICK00.PHVAST AS DISTI_FLAG,
      PHPICK00.PHI2O5 AS ORDER_HOLD,
      PHPICK00.PHSC6 AS CFI,
      CHCART00.CHSC1 AS PLT_FLAG,
      PHPICK00.PHI2O5 AS HOLD,
      PHPICK00.PHVAST AS BD_FLAG
    FROM
      CAPM01.WM0272PRDD.PHPICK00 PHPICK00
      JOIN CAPM01.WM0272PRDD.PDPICK00 PDPICK00 ON PHPICK00.PHPCTL = PDPICK00.PDPCTL
      JOIN CAPM01.WM0272PRDD.STSTYL00 STSTYL00 ON PDPICK00.PDSTYL = STSTYL00.STSTYL
      LEFT JOIN CAPM01.WM0272PRDD.CHCART00 CHCART00 ON PHPICK00.PHPCTL = CHCART00.CHPCTL
    WHERE
      PHPICK00.PHWHSE = 'BNA'
      AND STSTYL00.STDIV = '05'
      AND PHPICK00.PHPKTN <> ''
      AND PHPICK00.PHPSTF NOT IN ('00', '90', '95', '99')
  )
GROUP BY
  ORDER,
  STATUS,
  GLN,
  OT_FLAG,
  SVIA,
  TOTAL_QTY,
  WEIGHT,
  ESD,
  SCODE,
  DL_DATE_CST,
  COUNTRY,
  ORDER_TYPE,
  UPC_FLAG,
  DISTI_FLAG,
  ORDER_HOLD,
  CFI,
  GIT,
  PLT_FLAG,
  HOLD,
  BD_FLAG"
