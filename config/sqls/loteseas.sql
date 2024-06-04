SELECT 
	e.Nome as "UNIDADE",
	l.Id as "N° LOTE",
	l.`Data` as "DATA",
	CONCAT(l.Mes, "/",l.Ano) as "COMPETENCIA",
	SUM(IF(li.Erros IS NULL, 1, 0)) as "FICHAS VALIDAS",
	SUM(IF(li.Erros IS NOT NULL, 1, 0)) as "FICHAS INVALIDAS"
from lote l
inner join LoteIntegracao li on li.Lote_Id = l.Id
inner join Estabelecimento e on l.Estabelecimento_Id = e.Id 
WHERE 
	CONCAT(l.Mes,'/',l.Ano) = %(comp)s
group by 
	l.Id
order BY 
	l.Id desc