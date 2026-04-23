/****** Script do comando SelectTopNRows de SSMS  ******/

  
--update  [BHTGloboGrupo].[dbo].[tblBanco] set NomeBanco = 'Auto' where CodBanco = 1
--  GO
--  update [BHTGloboGrupo].[dbo].[tblBanco] set NomeBanco = 'Market Place', FlagAtivo = 1 where CodBanco = 2

--  GO


UPDATE BHTGloboAuto.dbo.tblTraducaoComponenteObjetoTraducao
SET Texto = CASE Texto
    WHEN 'Awareness - TOM' THEN 'Espontâneo - Top of Mind'
    WHEN 'Spontaneous - OM' THEN 'Espontâneo - Outras menções'
    WHEN 'Prompted' THEN 'Estimulado'
    WHEN 'Total Awareness' THEN 'Total Conhecimento'
END
WHERE Texto IN ('Awareness - TOM', 'Spontaneous - OM', 'Prompted', 'Total Awareness');

UPDATE BHTGloboAuto.dbo.tblTraducaoComponenteObjetoTraducao
SET Texto ='Conhecimento'
WHERE Objeto IN ('navbar.dashboard-two','sidebar.menu2')



UPDATE BHTGloboAuto.dbo.tblTraducaoComponenteObjetoTraducao
SET Texto ='Onda Anterior'
WHERE Objeto IN ('dashboard-three-leg-wave-anterior')


UPDATE BHTGloboMarketPlace.dbo.tblTraducaoComponenteObjetoTraducao
SET Texto = CASE Texto
    WHEN 'Awareness - TOM' THEN 'Espontâneo - Top of Mind'
    WHEN 'Spontaneous - OM' THEN 'Espontâneo - Outras menções'
    WHEN 'Prompted' THEN 'Estimulado'
    WHEN 'Total Awareness' THEN 'Total Conhecimento'
END
WHERE Texto IN ('Awareness - TOM', 'Spontaneous - OM', 'Prompted', 'Total Awareness');

UPDATE BHTGloboMarketPlace.dbo.tblTraducaoComponenteObjetoTraducao
SET Texto ='Conhecimento'
WHERE Objeto IN ('navbar.dashboard-two','sidebar.menu2')

UPDATE BHTGloboMarketPlace.dbo.tblTraducaoComponenteObjetoTraducao
SET Texto ='Onda Anterior'
WHERE Objeto IN ('dashboard-three-leg-wave-anterior')
