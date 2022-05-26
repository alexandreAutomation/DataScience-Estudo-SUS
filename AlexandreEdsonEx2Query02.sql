Select Qtd_alunos, Disciplina, Nota From 	   
(SELECT ta.nome Qtd_alunos, ta.ra, th.ra, th.cod_disc, td.cod_disc, td.nome_disc Disciplina, th.nota Nota
FROM tabela_alunos as ta
INNER JOIN tabela_historico as th
on ta.ra = th.ra
INNER JOIN tabela_disciplinas as td
on th.cod_disc = td.cod_disc

where nota >= 6
)