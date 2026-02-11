# Controle de Consumo de Tinta (Excel 365, unidade: Quilo)

Abas/Tabelas no .xlsx:
- Cores (Tabela: T_Cores): Cor
- Entradas (Tabela: T_Entradas): Data | Cor | Qtd (kg)
- Estoque_Mensal (Tabela: T_EstoqueMensal): Mês | Cor | Estoque Final (kg) | Estoque Inicial (kg) (opcional)
- Resumo (Tabela: T_Resumo): Mês | Cor | Inicial (kg) | Entradas (kg) | Final (kg) | Consumo (kg)

Fórmulas (colunas da planilha Resumo):
- Inicial (kg):
=IFERROR(INDEX(Estoque_Mensal!$C:$C,SUMPRODUCT((Estoque_Mensal!$B:$B=B2)*(Estoque_Mensal!$A:$A=EOMONTH(A2,-1))*ROW(Estoque_Mensal!$B:$B))),IFERROR(INDEX(Estoque_Mensal!$D:$D,SUMPRODUCT((Estoque_Mensal!$B:$B=B2)*(Estoque_Mensal!$A:$A=A2)*ROW(Estoque_Mensal!$B:$B))),""))

- Entradas (kg):
=IFERROR(SUMIFS(Entradas!$C:$C,Entradas!$B:$B,B2,Entradas!$A:$A,">="&DATE(YEAR(A2),MONTH(A2),1),Entradas!$A:$A,"<="&EOMONTH(A2,0)),0)

- Final (kg):
=IFERROR(INDEX(Estoque_Mensal!$C:$C,SUMPRODUCT((Estoque_Mensal!$B:$B=B2)*(Estoque_Mensal!$A:$A=A2)*ROW(Estoque_Mensal!$B:$B))),"")

- Consumo (kg):
=C2 + D2 - E2

Uso:
1) Lance Entradas durante o mês.
2) No último dia, lance Estoque Final por cor (e o Inicial só no 1º mês de cada cor).
3) A planilha Resumo calcula automaticamente o Consumo por mês/cor.
