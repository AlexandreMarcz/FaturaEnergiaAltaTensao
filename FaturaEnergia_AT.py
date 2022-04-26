#!/usr/bin/env python
# -*- coding: utf-8 -*-


import pandas
import matplotlib.pyplot as plt

arquivo_leitura =  pandas.read_excel('Dados.xlsx', sheet_name='Planilha1')

# IMPOSTOS A SEREM CONSIDERADOS

ICMS = 0.29 
PIS = 0.0088
COFINS = 0.0404

fatorPC = 1/(1-PIS-COFINS)
fatorICMS = 1/(1-ICMS)

# TARIFAS A4 Copel Dis. Caso o cliente nao seja A4, necessario atualizar as tarifas
tarifa_demanda_verde = 17.07   
tarifa_demanda_pontaazul = 36.08 
tarifa_demanda_forapontaazul = 17.07 
tarifa_tusd_pontaverde = 959.22 
tarifa_tusd_forapontverde = 82.6 
tarifa_tusd_pontaazul = 82.6 
tarifa_tusd_forapontaazul = 82.6 

tarifa_energia_ponta =  437.87 
tarifa_energia_foraponta = 275.36 

#BANDEIRAS TARIFARIAS vigentes

amarela = 18.74

vermelha1 = 39.71

vermelha2 = 94.92

escassez_hidrica = 142

# Abaixo segue o esqueleto do calculo da fatura. Essa classe 'Conta' refere-se a
#conta de energia do cliente que sera gerada conforme o consumo dele

class Conta():
	
	def __init__(self,eponta, eforaponta, dponta, dforaponta,dcponta,dcforaponta):
		self.eponta = eponta
		self.eforaponta = eforaponta
		self.dponta = dponta
		self.dforaponta = dforaponta
		self.dcponta = dcponta
		self.dcforaponta = dcforaponta
		
		
	def energia_cativo(self):
		valorep = self.eponta*0.001*tarifa_energia_ponta*fatorPC
		valorep = valorep*fatorICMS
		
		valorefp = self.eforaponta*0.001*tarifa_energia_foraponta*fatorPC
		valorefp = valorefp*fatorICMS
		
		return valorep + valorefp
		
			
	
	def demanda_verde(self):
		
		valor_dem = max(self.dponta,self.dforaponta)*tarifa_demanda_verde*fatorPC
		valor_dem = valor_dem*fatorICMS
		
		if max(self.dcponta,self.dcforaponta)<= max(self.dponta,self.dforaponta):
			dem_isenta = 0
		else:
			dem_isenta = (max(self.dcponta,self.dcforaponta)-max(self.dponta,self.dforaponta))*tarifa_demanda_verde*fatorPC
		
		return valor_dem + dem_isenta
	
	
	def demanda_azul(self):
		
		valor_dem_p = self.dponta*tarifa_demanda_pontaazul*fatorPC
		valor_dem_p = valor_dem_p*fatorICMS
		
		valor_dem_fp = self.dforaponta*tarifa_demanda_forapontaazul*fatorPC
		valor_dem_fp = valor_dem_fp*fatorICMS
		
		if self.dcponta-self.dponta >0:
			dem_isenta_p = (self.dcponta-self.dponta)*tarifa_demanda_pontaazul*fatorPC
		else:
			dem_isenta_p = 0
		
		if self.dcforaponta-self.dforaponta > 0:
			dem_isenta_fp = (self.dcforaponta-self.dforaponta)*tarifa_demanda_forapontaazul*fatorPC
		else:
			dem_isenta_fp = 0
			
		return valor_dem_p + valor_dem_fp + dem_isenta_p + dem_isenta_fp
		
		
		
	def tusd_verde(self):
		valor_tusdP = self.eponta*0.001*tarifa_tusd_pontaverde*fatorPC
		valor_tusdP = valor_tusdP*fatorICMS
		
		valor_tusdFP = self.eforaponta*0.001*tarifa_tusd_forapontverde*fatorPC
		valor_tusdFP = valor_tusdFP*fatorICMS
		
		return valor_tusdP + valor_tusdFP
		
	def tusd_azul(self):
		valor_tusdP = self.eponta*0.001*tarifa_tusd_pontaazul*fatorPC
		valor_tusdP = valor_tusdP*fatorICMS
		
		valor_tusdFP = self.eforaponta*0.001*tarifa_tusd_forapontaazul*fatorPC
		valor_tusdFP = valor_tusdFP*fatorICMS
		
		return valor_tusdP + valor_tusdFP
		
	
	def bandeira_amarela(self):
		valor_band = (self.eponta+self.eforaponta)*0.001*amarela*fatorPC
		valor_band = valor_band*fatorICMS
		
		return valor_band
		
	def bandeira_vermelha1(self):
		valor_band = (self.eponta+self.eforaponta)*0.001*vermelha1*fatorPC
		valor_band = valor_band*fatorICMS
		
		return valor_band
	
	def bandeira_vermelha2(self):
		valor_band = (self.eponta+self.eforaponta)*0.001*vermelha2*fatorPC
		valor_band = valor_band*fatorICMS
		
		return valor_band
		
	def bandeira_escassez(self):
		valor_band = (self.eponta+self.eforaponta)*0.001*escassez_hidrica*fatorPC
		valor_band = valor_band*fatorICMS
		
		return valor_band
	

# A rotina abaixo vai fazer a leitura e gerar os dados da conta do cliente

for x in range(0,len(arquivo_leitura.index)): # len(arquivo_leitura.index)
    
    
    a = arquivo_leitura.loc[x,'EP'] #Consumo Ponta
    b = arquivo_leitura.loc[x,'EFP'] #Consumo Fora Ponta
    c = arquivo_leitura.loc[x,'DCP'] # Demanda Contratada Ponta
    d = arquivo_leitura.loc[x,'DCFP'] # Demanda contratada Fora Ponta
    e = arquivo_leitura.loc[x,'DMP'] # Demanda na Ponta
    f = arquivo_leitura.loc[x,'DMFP'] # Demanda Fora de Ponta
    
    resultado = Conta(a,b,e,f,c,d)

    arquivo_leitura.loc[x,'Cativo Verde'] = resultado.energia_cativo() + resultado.demanda_verde() + resultado.tusd_verde()
    
    arquivo_leitura.loc[x,'Cativo Azul'] = resultado.energia_cativo() + resultado.demanda_azul() + resultado.tusd_azul()
    
    arquivo_leitura.loc[x,'Cativo Verde B Amarela'] = resultado.energia_cativo() + resultado.demanda_verde() + resultado.tusd_verde() + resultado.bandeira_amarela()
    
    arquivo_leitura.loc[x,'Cativo Verde Vermelha 1'] = resultado.energia_cativo() + resultado.demanda_verde() + resultado.tusd_verde() + resultado.bandeira_vermelha1()
    
    arquivo_leitura.loc[x,'Cativo Verde Vermelha 2'] = resultado.energia_cativo() + resultado.demanda_verde() + resultado.tusd_verde() + resultado.bandeira_vermelha2()
    
    arquivo_leitura.loc[x,'Cativo Verde Escassez'] = resultado.energia_cativo() + resultado.demanda_verde() + resultado.tusd_verde() + resultado.bandeira_escassez()
    
    arquivo_leitura.loc[x,'Cativo Azul B Amarela'] = resultado.energia_cativo() + resultado.demanda_azul() + resultado.tusd_azul() + resultado.bandeira_amarela()
    
    arquivo_leitura.loc[x,'Cativo Azul Vermelha 1'] = resultado.energia_cativo() + resultado.demanda_azul() + resultado.tusd_azul() + resultado.bandeira_vermelha1()
    
    arquivo_leitura.loc[x,'Cativo Azul Vermelha 2'] = resultado.energia_cativo() + resultado.demanda_azul() + resultado.tusd_azul() + resultado.bandeira_vermelha2()
    
    arquivo_leitura.loc[x,'Cativo Azul Escassez'] = resultado.energia_cativo() + resultado.demanda_azul() + resultado.tusd_azul() + resultado.bandeira_escassez()
    

    print('estamos no passo '+str(x))


arquivo_leitura['Somatorio Cativo Verde'] = arquivo_leitura['Cativo Verde'].cumsum() 

arquivo_leitura['Somatorio Cativo Azul'] = arquivo_leitura['Cativo Azul'].cumsum()    


# Linha para exportar os dados gerados para uma planilha
arquivo_leitura.to_excel('ConsumoEnergia_A4.xlsx',sheet_name = 'Planilha1', index=False)


# Parte para plotar os dados gerados

distancia = arquivo_leitura.loc[len(arquivo_leitura.index)-1,'Somatorio Cativo Azul'] - arquivo_leitura.loc[len(arquivo_leitura.index)-1,'Somatorio Cativo Verde']

d_distancia = round(distancia,2)

plt.plot(arquivo_leitura['Data'],arquivo_leitura['Somatorio Cativo Verde'],color='green' )
plt.plot(arquivo_leitura['Data'],arquivo_leitura['Somatorio Cativo Azul'],color='blue')
plt.xlabel('Data')
plt.ylabel('R$')
plt.legend(['Cativo Verde','Cativo Azul'])
plt.title('Acumulado - Comparativo Modalidade Tarifa Horaria Verde e Azul')
plt.text(arquivo_leitura.loc[len(arquivo_leitura.index)-1,'Data'],arquivo_leitura.loc[len(arquivo_leitura.index)-1,'Somatorio Cativo Verde']+distancia/2," R$ "+str(d_distancia))
plt.vlines(arquivo_leitura.loc[len(arquivo_leitura.index)-1,'Data'],arquivo_leitura.loc[len(arquivo_leitura.index)-1,'Somatorio Cativo Verde'],arquivo_leitura.loc[len(arquivo_leitura.index)-1,'Somatorio Cativo Azul'], color = 'red')
plt.show()

plt.clf()


plt.plot(arquivo_leitura['Data'],arquivo_leitura['Cativo Verde'], color = 'green')
plt.plot(arquivo_leitura['Data'],arquivo_leitura['Cativo Verde B Amarela'],color='yellow', )
plt.plot(arquivo_leitura['Data'],arquivo_leitura['Cativo Verde Vermelha 1'], color = 'red',alpha=0.5 )
plt.plot(arquivo_leitura['Data'],arquivo_leitura['Cativo Verde Vermelha 2'], color = 'red')
plt.plot(arquivo_leitura['Data'],arquivo_leitura['Cativo Verde Escassez'], color = 'purple')
plt.legend(['Bandeira Verde R$ 0/MWh','Bandeira Amarela R$'+str(amarela)+'/MWh','Bandeira Vermelha 1 R$'+str(vermelha1)+'/MWh','Bandeira Vermelha 2 R$'+str(vermelha2)+'/MWh','Bandeira Escassez Hidrica R$'+str(escassez_hidrica)+'/MWh'])
plt.title('Custo Mensal Bandeiras - Tarifa Horaria Verde')
plt.xlabel('Data')
plt.ylabel('R$')

plt.show()