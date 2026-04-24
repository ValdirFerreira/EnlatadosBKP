-- =============================================
-- Traduções: Brand Funil (IdiomaId = 1 - Português)
-- Conforme imagem:
--   Conhecimento   ? já correto (id 75)
--   Consideração   ? já correto (id 76)
--   Calc6          ? Últimos 6 meses (NOVO)
--   Calc7          ? Últimos 3 meses (NOVO)
--   Uso            ? Último Mês     (era "Posse")
--   Preferencia    ? Preferência    (era "Primeira escolha")
--   Loyalty        ? Rejeição       ? já correto (id 120)
-- =============================================

USE [BHTEnlatados]
GO

-- -----------------------------------------------
-- grafico-funil-uso: "Posse" ? "Último Mês"
-- -----------------------------------------------
IF EXISTS (
    SELECT 1 FROM [dbo].[tblTraducaoComponenteObjetoTraducao]
    WHERE [Objeto] = 'grafico-funil-uso' AND [IdiomaId] = 1
)
    UPDATE [dbo].[tblTraducaoComponenteObjetoTraducao]
    SET [Texto] = N'Último Mês'
    WHERE [Objeto] = 'grafico-funil-uso' AND [IdiomaId] = 1;
ELSE
    INSERT INTO [dbo].[tblTraducaoComponenteObjetoTraducao] ([IdiomaId], [Objeto], [Texto])
    VALUES (1, N'grafico-funil-uso', N'Último Mês');
GO

-- -----------------------------------------------
-- grafico-funil-preferencia: "Primeira escolha" ? "Preferência"
-- -----------------------------------------------
IF EXISTS (
    SELECT 1 FROM [dbo].[tblTraducaoComponenteObjetoTraducao]
    WHERE [Objeto] = 'grafico-funil-preferencia' AND [IdiomaId] = 1
)
    UPDATE [dbo].[tblTraducaoComponenteObjetoTraducao]
    SET [Texto] = N'Preferência'
    WHERE [Objeto] = 'grafico-funil-preferencia' AND [IdiomaId] = 1;
ELSE
    INSERT INTO [dbo].[tblTraducaoComponenteObjetoTraducao] ([IdiomaId], [Objeto], [Texto])
    VALUES (1, N'grafico-funil-preferencia', N'Preferência');
GO

-- -----------------------------------------------
-- grafico-funil-Calc6: NOVO ? "Últimos 6 meses"
-- -----------------------------------------------
IF EXISTS (
    SELECT 1 FROM [dbo].[tblTraducaoComponenteObjetoTraducao]
    WHERE [Objeto] = 'grafico-funil-Calc6' AND [IdiomaId] = 1
)
    UPDATE [dbo].[tblTraducaoComponenteObjetoTraducao]
    SET [Texto] = N'Últimos 6 meses'
    WHERE [Objeto] = 'grafico-funil-Calc6' AND [IdiomaId] = 1;
ELSE
    INSERT INTO [dbo].[tblTraducaoComponenteObjetoTraducao] ([IdiomaId], [Objeto], [Texto])
    VALUES (1, N'grafico-funil-Calc6', N'Últimos 6 meses');
GO

-- -----------------------------------------------
-- grafico-funil-Calc7: NOVO ? "Últimos 3 meses"
-- -----------------------------------------------
IF EXISTS (
    SELECT 1 FROM [dbo].[tblTraducaoComponenteObjetoTraducao]
    WHERE [Objeto] = 'grafico-funil-Calc7' AND [IdiomaId] = 1
)
    UPDATE [dbo].[tblTraducaoComponenteObjetoTraducao]
    SET [Texto] = N'Últimos 3 meses'
    WHERE [Objeto] = 'grafico-funil-Calc7' AND [IdiomaId] = 1;
ELSE
    INSERT INTO [dbo].[tblTraducaoComponenteObjetoTraducao] ([IdiomaId], [Objeto], [Texto])
    VALUES (1, N'grafico-funil-Calc7', N'Últimos 3 meses');
GO

-- -----------------------------------------------
-- Verificação final
-- -----------------------------------------------
SELECT [Id], [IdiomaId], [Objeto], [Texto]
FROM [dbo].[tblTraducaoComponenteObjetoTraducao]
WHERE [Objeto] IN (
    'grafico-funil-conhecimento',
    'grafico-funil-consideracao',
    'grafico-funil-uso',
    'grafico-funil-preferencia',
    'grafico-funil-Loyalty',
    'grafico-funil-Calc6',
    'grafico-funil-Calc7'
)
AND [IdiomaId] = 1
ORDER BY [Id];
GO