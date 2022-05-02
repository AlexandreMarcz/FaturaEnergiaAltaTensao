# FaturaEnergiaAltaTensao
Script para simular conta de energia de cliente grupo A Alta Tensão
Para gerar a simulação é necessário criar uma planilha contendo as seguintes colunas:

*Data - data em que houveram as medições;

*EP - Energia consumida no horário de  Ponta, em KWh;

*EFP - Energia consumida no horário Fora Ponta, em KWh;

*DCP - Demanda Contratada no horário Ponta, em KW;

*DCFP - Demanda Contratada no horário Fora Ponta, em KW;

*DMP - Demanda Medida no horário Ponta, em KW;

*DMFP - Demanda Medida no horário Fora Ponta, em KW;



Importante: Se o cliente que for ser simulado for do tipo horário verde, os dados de DEMANDA PONTA e FORA PONTA devem ser iguais na planilha, tanto a Medida quanto a Contratada. A planilha gerada deve ficar na mesma pasta que este script.

Atentar no script quanto a adequação das aliquotas de imposto e de tarifas, conforme a distribuidora de energia a ser considerada.
