SET LANGUAGE Portuguese;

WITH Batidas AS (
    SELECT 
        HF.CHAPA,
        HF.DATA,
        B.BATIDA,
        ROW_NUMBER() OVER (PARTITION BY HF.CHAPA, HF.DATA ORDER BY B.BATIDA) AS RN
    FROM 
        AAFHTFUN HF
        LEFT JOIN ABATFUN B ON B.CODCOLIGADA = HF.CODCOLIGADA AND B.CHAPA = HF.CHAPA AND B.DATA = HF.DATA
    WHERE 
        HF.CODCOLIGADA = :CODCOLIGADA
        AND HF.DATA BETWEEN :Data_Inicio AND :Data_Fim
),
Abonos AS (
    SELECT 
        AF.CHAPA,
        AF.DATA,
        SUM(ISNULL(AF.HORAFIM - AF.HORAINICIO, 0)) AS [TOTAL_MINUTOS_ABONO]
    FROM 
        AABONFUN AF
    WHERE 
        AF.CODCOLIGADA = :CODCOLIGADA
        AND AF.DATA BETWEEN :Data_Inicio AND :Data_Fim
    GROUP BY 
        AF.CHAPA, AF.DATA
),
Entradas AS (
    SELECT 
        HF.CHAPA,
        HF.DATA,
        CASE 
            WHEN SUM(CASE WHEN B.RN = 1 AND B.BATIDA > 0 THEN 1 ELSE 0 END) = 0 
                THEN NULL
            ELSE
                RIGHT('0' + CAST((SUM(CASE WHEN B.RN = 1 THEN B.BATIDA ELSE 0 END))/60 AS VARCHAR), 2) + ':' +
                RIGHT('0' + CAST((SUM(CASE WHEN B.RN = 1 THEN B.BATIDA ELSE 0 END)) % 60 AS VARCHAR), 2)
        END AS ENTRADA
    FROM AAFHTFUN HF
    LEFT JOIN Batidas B 
        ON B.CHAPA = HF.CHAPA AND B.DATA = HF.DATA
    WHERE HF.CODCOLIGADA = :CODCOLIGADA
        AND HF.DATA BETWEEN :Data_Inicio AND :Data_Fim
    GROUP BY HF.CHAPA, HF.DATA
),
SequenciaBase AS (
    SELECT 
        E.CHAPA,
        E.DATA,
        E.ENTRADA,
        LAG(E.ENTRADA) OVER (PARTITION BY E.CHAPA ORDER BY E.DATA) AS ENTRADA_ANT,
        LAG(E.DATA) OVER (PARTITION BY E.CHAPA ORDER BY E.DATA) AS DATA_ANT
    FROM Entradas E
),
Quebras AS (
    SELECT *,
        CASE 
            WHEN ENTRADA IS NULL THEN NULL
            WHEN ENTRADA_ANT IS NULL THEN 1
            WHEN DATEDIFF(DAY, DATA_ANT, DATA) > 1 THEN 1
            WHEN ENTRADA_ANT IS NULL THEN 1
            ELSE 0
        END AS NOVA_SEQ
    FROM SequenciaBase
),
Grupos AS (
    SELECT *,
        SUM(CASE WHEN NOVA_SEQ = 1 THEN 1 ELSE 0 END)
        OVER (PARTITION BY CHAPA ORDER BY DATA ROWS UNBOUNDED PRECEDING) AS GRUPO_SEQ
    FROM Quebras
),
GruposComInicio As (    
	SELECT
    	SB.CHAPA,
        SB.DATA,
        SB.ENTRADA,
        G.GRUPO_SEQ,
        CASE 
            WHEN SB.ENTRADA IS NOT NULL THEN
                MIN(SB.DATA) OVER (PARTITION BY SB.CHAPA, G.GRUPO_SEQ)
            ELSE NULL
        END AS DATA_INICIO_SEQ
    FROM SequenciaBase SB
    LEFT JOIN Grupos G ON SB.CHAPA = G.CHAPA AND SB.DATA = G.DATA
),
SequenciaFinal AS (
    SELECT 
        CHAPA,
        DATA,
        CASE 
            WHEN ENTRADA IS NOT NULL THEN 
                ROW_NUMBER() OVER (PARTITION BY CHAPA, GRUPO_SEQ ORDER BY DATA)
        END AS SEQUENCIA
    FROM Grupos
),
SequenciaTotalFinal AS (
    SELECT 
        CHAPA,
        DATA,
        COUNT(*) OVER (PARTITION BY CHAPA, GRUPO_SEQ) AS TAMANHO_SEQ,
        CASE 
            WHEN ROW_NUMBER() OVER (PARTITION BY CHAPA, GRUPO_SEQ ORDER BY DATA DESC) = 1
            THEN 'SEQ - ' + CAST(COUNT(*) OVER (PARTITION BY CHAPA, GRUPO_SEQ) AS VARCHAR)
        END AS SEQUENCIA_TOTAL
    FROM Grupos
    WHERE ENTRADA IS NOT NULL
),
-- Gerando pseudônimo para CHAPA (exemplo usando HASH, depende do SGBD)
Anonimizados AS (
    SELECT DISTINCT
        CHAPA,
        -- Exemplo genérico: concatena prefixo + chapa para anonimizar
        'COLAB_' + CAST(ROW_NUMBER() OVER (ORDER BY CHAPA) AS VARCHAR) AS ANON_CHAPA
    FROM AAFHTFUN
    WHERE CODCOLIGADA = :CODCOLIGADA
)

SELECT 
     F.CODCOLIGADA 																								AS [COLIGADA]
    ,A.ANON_CHAPA																								AS [CHAPA_ANONIMIZADA]
    ,'Colaborador ' + CAST(ROW_NUMBER() OVER (PARTITION BY F.CODCOLIGADA ORDER BY F.CODSECAO, F.CODSITUACAO) AS VARCHAR) AS [COLABORADOR]
    ,HF.DATA																									AS [DT.APURACAO]
    ,GCI.DATA_INICIO_SEQ																						AS [PERÍODO]
    ,DATENAME(WEEKDAY, HF.DATA)																					AS [DIA SEMANA]
    ,F.CODSECAO 																								AS [SECAO]
    ,S.DESCRICAO 																								AS [PROJETO]
    ,NULL																										AS [MAO DE OBRA]
    ,F.CODSITUACAO 																								AS [SITUACAO]
    ,NULL																										AS [DATA ADMISSAO]
    ,NULL																										AS [DATA RESCISAO]
    ,'Cargo X'																									AS [DESCR. CARGO]

    ,CASE
        WHEN SUM(CASE WHEN B.RN = 1 AND B.BATIDA > 0 THEN 1 ELSE 0 END) = 0 
        THEN NULL
        ELSE
            RIGHT('0' + CAST((SUM(CASE WHEN B.RN = 1 THEN B.BATIDA ELSE 0 END))/60 AS VARCHAR), 2) + ':' +
            RIGHT('0' + CAST((SUM(CASE WHEN B.RN = 1 THEN B.BATIDA ELSE 0 END)) % 60 AS VARCHAR), 2)
    	END 																									AS [ENTRADA]
	,SF.SEQUENCIA																								AS [SEQUENCIA]
	,ST.SEQUENCIA_TOTAL																							AS [SEQUENCIATOTAL]
	
    ,CASE 
        WHEN SUM(CASE WHEN B.RN = 2 AND B.BATIDA > 0 THEN 1 ELSE 0 END) = 0 
        THEN NULL
        ELSE
            RIGHT('0' + CAST((SUM(CASE WHEN B.RN = 2 THEN B.BATIDA ELSE 0 END))/60 AS VARCHAR), 2) + ':' +
            RIGHT('0' + CAST((SUM(CASE WHEN B.RN = 2 THEN B.BATIDA ELSE 0 END)) % 60 AS VARCHAR), 2)
    	END 																									AS [SAIDA]
    ,CASE 
        WHEN SUM(CASE WHEN B.RN = 3 AND B.BATIDA > 0 THEN 1 ELSE 0 END) = 0 
        THEN NULL
        ELSE
            RIGHT('0' + CAST((SUM(CASE WHEN B.RN = 3 THEN B.BATIDA ELSE 0 END))/60 AS VARCHAR), 2) + ':' +
            RIGHT('0' + CAST((SUM(CASE WHEN B.RN = 3 THEN B.BATIDA ELSE 0 END)) % 60 AS VARCHAR), 2)
    END 																										AS [ENTRADA1]
    ,CASE 
        WHEN SUM(CASE WHEN B.RN = 4 AND B.BATIDA > 0 THEN 1 ELSE 0 END) = 0 
        THEN NULL
        ELSE
            RIGHT('0' + CAST((SUM(CASE WHEN B.RN = 4 THEN B.BATIDA ELSE 0 END))/60 AS VARCHAR), 2) + ':' +
            RIGHT('0' + CAST((SUM(CASE WHEN B.RN = 4 THEN B.BATIDA ELSE 0 END)) % 60 AS VARCHAR), 2)
    END 																										AS [SAIDA1]
    
    ,CASE 
    	WHEN ST.SEQUENCIA_TOTAL IS NULL THEN NULL
    	WHEN ST.TAMANHO_SEQ BETWEEN 1 AND 6 THEN 'OK'
    	WHEN ST.TAMANHO_SEQ BETWEEN 7 AND 14 THEN 'Alerta'
    	WHEN ST.TAMANHO_SEQ >= 15 THEN 'Crítico'																								
    END 																									AS [CLASSIFICACAO]
    
FROM 
    AAFHTFUN HF
    
    JOIN PFUNC F
        ON F.CODCOLIGADA = HF.CODCOLIGADA 
        AND F.CHAPA = HF.CHAPA
    
    JOIN PSECAO S 
        ON S.CODCOLIGADA = F.CODCOLIGADA 
        AND S.CODIGO = F.CODSECAO
    
    LEFT JOIN PFUNCAO FN 
        ON FN.CODIGO = F.CODFUNCAO 
        AND FN.CODCOLIGADA = F.CODCOLIGADA
    
    LEFT JOIN Abonos Abo 
        ON Abo.CHAPA = HF.CHAPA 
        AND Abo.DATA = HF.DATA
    
    LEFT JOIN Batidas B 
        ON B.CHAPA = HF.CHAPA 
        AND B.DATA = HF.DATA
	
	LEFT JOIN SequenciaFinal SF
        ON SF.CHAPA = HF.CHAPA
        AND SF.DATA = HF.DATA
	
	LEFT JOIN SequenciaTotalFinal ST
        ON ST.CHAPA = HF.CHAPA
        AND ST.DATA = HF.DATA
	
	LEFT JOIN GruposComInicio GCI
        ON GCI.CHAPA = HF.CHAPA
        AND GCI.DATA = HF.DATA

    JOIN Anonimizados A
        ON A.CHAPA = HF.CHAPA
    
WHERE 
    HF.CODCOLIGADA = :CODCOLIGADA
    AND HF.DATA BETWEEN :Data_Inicio AND :Data_Fim

ORDER BY 
    HF.DATA ASC