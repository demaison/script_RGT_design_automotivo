# instala o pacote OpenRepGrid
install.packages("OpenRepGrid")

# instala o pacote corrplot
install.packages("corrplot")

# instala o pacote writexl
install.packages("writexl")

# carrega os pacotes
library(OpenRepGrid)
library(corrplot)
library(writexl)

# indica diretório atual
getwd()

# importa planilha de excel para o dataframe rgt_grid e transforma em um grid
rgt_grid_esp <- importExcel("planilhaRGT_esp.xlsx")
rgt_grid_esp_semacento <- importExcel("planilhaRGT_esp_semacento.xlsx")

rgt_grid_nesp <- importExcel("planilhaRGT_nesp.xlsx")
rgt_grid_nesp_semacento <- importExcel("planilhaRGT_nesp_semacento.xlsx")

# faz a análise de componentes principais (PCA), com padrão de rotação Varimax e reduzindo a duas dimensões
constructPca(rgt_grid_esp, nf=2)
constructPca(rgt_grid_esp_semacento, nf=2)
constructPca(rgt_grid_nesp, nf=2)
constructPca(rgt_grid_nesp_semacento, nf=2)

# gera gráficos Biplot 3D e Biplot simples
biplot3d(rgt_grid_esp_semacento, labels.c = TRUE, e.cex = 1, c.text.col = "gray49", e.text.col = "darkblue")
biplot3d(rgt_grid_nesp_semacento, labels.c = TRUE, e.cex = 1, c.text.col = "gray49", e.text.col = "darkred")
biplotSimple(rgt_grid_esp, c.label.col="darkblue", e.label.cex = 1, zoom = 3.5)
biplotSimple(rgt_grid_nesp, c.label.col="darkred", e.label.cex = 1, zoom = 3.5)

# gera matrizes ordenáveis de Bertin e plota em pedaços (a cada 25 linhas)
# obs.: células claras correspondem a pontuações baixas, escuras a altas
bertin(rgt_grid_esp[1:25, 1:5], color=c("white", "darkblue"), xlim = c(0.3,0.4), ylim = c(.0, .8), lheight = 0.8,id = c(F, F))
bertin(rgt_grid_esp[26:50, 1:5], color=c("white", "darkblue"), xlim = c(0.3,0.4), ylim = c(.0, .8), lheight = 0.8,id = c(F, F))
bertin(rgt_grid_esp[51:75, 1:5], color=c("white", "darkblue"), xlim = c(0.3,0.4), ylim = c(.0, .8), lheight = 0.8,id = c(F, F))
bertin(rgt_grid_esp[76:100, 1:5], color=c("white", "darkblue"), xlim = c(0.3,0.4), ylim = c(.0, .8), lheight = 0.8,id = c(F, F))
bertin(rgt_grid_nesp[1:25, 1:5], color=c("white", "darkred"), xlim = c(0.3,0.4), ylim = c(.0, .8), lheight = 0.8,id = c(F, F))
bertin(rgt_grid_nesp[26:50, 1:5], color=c("white", "darkred"), xlim = c(0.3,0.4), ylim = c(.0, .8), lheight = 0.8,id = c(F, F))
bertin(rgt_grid_nesp[51:75, 1:5], color=c("white", "darkred"), xlim = c(0.3,0.4), ylim = c(.0, .8), lheight = 0.8,id = c(F, F))
bertin(rgt_grid_nesp[76:100, 1:5], color=c("white", "darkred"), xlim = c(0.3,0.4), ylim = c(.0, .8), lheight = 0.8,id = c(F, F))

# clusteriza construtos e os veículos para gerar dendogramas
cluster(rgt_grid_esp, xlim = c(40,0)) 
cluster(rgt_grid_nesp, xlim = c(40,0))

# salva as estatísticas básicas do RGT em um dataframe
estatisticas_esp <- statsConstructs(rgt_grid_esp)
estatisticas_nesp <- statsConstructs(rgt_grid_nesp)

# gera uma matriz de correlação e salva em um dataframe, usando coeficiente de Pearson
# obs.: útil para visualizar rapidamente correlações mais fortes
matriz_correlacao_esp <- constructCor(rgt_grid_esp)
matriz_correlacao_nesp <- constructCor(rgt_grid_nesp)

# cria corrplot para visualizar a matriz de correlação
# obs.: o conjunto de dados é muito grande para ser corretamente visualizado
corrplot(matriz_correlacao_esp, method ='number')       
corrplot(matriz_correlacao_nesp, method ='number')

# exporta um csv com a matriz de correlação desejada (especialistas e não-especialistas)
write.csv(matriz_correlacao_esp, "matriz_especialistas.csv")
write.csv(matriz_correlacao_nesp, "matriz_nao_especialistas.csv")

# passo a passo para criar as visualizações individuais para os especialistas
# extrai subconjuntos do RGT
especialista1<-rgt_grid_esp[1:10, 1:5]    
especialista2<-rgt_grid_esp[11:20, 1:5]
especialista4<-rgt_grid_esp[31:40, 1:5]
especialista5<-rgt_grid_esp[41:50, 1:5]
especialista6<-rgt_grid_esp[51:60, 1:5]
especialista7<-rgt_grid_esp[61:70, 1:5]
especialista8<-rgt_grid_esp[71:80, 1:5]
especialista9<-rgt_grid_esp[81:90, 1:5]
especialista10<-rgt_grid_esp[91:100, 1:5]

# cria as matrizes de correlação individuais
matriz_correlacao_esp1 <- constructCor(especialista1)
matriz_correlacao_esp2 <- constructCor(especialista2)
matriz_correlacao_esp3 <- constructCor(especialista3) 
matriz_correlacao_esp4 <- constructCor(especialista4) 
matriz_correlacao_esp5 <- constructCor(especialista5) 
matriz_correlacao_esp6 <- constructCor(especialista6) 
matriz_correlacao_esp7 <- constructCor(especialista7) 
matriz_correlacao_esp8 <- constructCor(especialista8) 
matriz_correlacao_esp9 <- constructCor(especialista9) 
matriz_correlacao_esp10 <- constructCor(especialista10) 

# plota as matrizes de correlação individuais
corrplot(matriz_correlacao_esp1, method ='number', type = "lower", number.cex = 0.8, tl.pos = "n", cl.cex = 0.8)  
corrplot(matriz_correlacao_esp2, method ='number', type = "lower", number.cex = 0.8, tl.pos = "n", cl.cex = 0.8)
corrplot(matriz_correlacao_esp3, method ='number', type = "lower", number.cex = 0.8, tl.pos = "n", cl.cex = 0.8)
corrplot(matriz_correlacao_esp4, method ='number', type = "lower", number.cex = 0.8, tl.pos = "n", cl.cex = 0.8)
corrplot(matriz_correlacao_esp5, method ='number', type = "lower", number.cex = 0.8, tl.pos = "n", cl.cex = 0.8)
corrplot(matriz_correlacao_esp6, method ='number', type = "lower", number.cex = 0.8, tl.pos = "n", cl.cex = 0.8)
corrplot(matriz_correlacao_esp7, method ='number', type = "lower", number.cex = 0.8, tl.pos = "n", cl.cex = 0.8)
corrplot(matriz_correlacao_esp8, method ='number', type = "lower", number.cex = 0.8, tl.pos = "n", cl.cex = 0.8)
corrplot(matriz_correlacao_esp9, method ='number', type = "lower", number.cex = 0.8, tl.pos = "n", cl.cex = 0.8)
corrplot(matriz_correlacao_esp10, method ='number', type = "lower", number.cex = 0.8, tl.pos = "n", cl.cex = 0.8)

# gera matrizes ordenáveis de Bertin individuais
bertin(especialista1, color=c("white", "darkblue"), xlim = c(0.3,0.4), ylim = c(.0, 0.5), lheight = 0.7,id = c(F, F), margins = c(0,0.7,0.7)) 
bertin(especialista2, color=c("white", "darkblue"), xlim = c(0.3,0.4), ylim = c(.0, 0.5), lheight = 0.7,id = c(F, F), margins = c(0,0.7,0.7))
bertin(especialista3, color=c("white", "darkblue"), xlim = c(0.3,0.4), ylim = c(.0, 0.5), lheight = 0.7,id = c(F, F), margins = c(0,0.7,0.7))
bertin(especialista4, color=c("white", "darkblue"), xlim = c(0.3,0.4), ylim = c(.0, 0.5), lheight = 0.7,id = c(F, F), margins = c(0,0.7,0.7))
bertin(especialista5, color=c("white", "darkblue"), xlim = c(0.3,0.4), ylim = c(.0, 0.5), lheight = 0.7,id = c(F, F), margins = c(0,0.7,0.7))
bertin(especialista6, color=c("white", "darkblue"), xlim = c(0.3,0.4), ylim = c(.0, 0.5), lheight = 0.7,id = c(F, F), margins = c(0,0.7,0.7))
bertin(especialista7, color=c("white", "darkblue"), xlim = c(0.3,0.4), ylim = c(.0, 0.5), lheight = 0.7,id = c(F, F), margins = c(0,0.7,0.7))
bertin(especialista8, color=c("white", "darkblue"), xlim = c(0.3,0.4), ylim = c(.0, 0.5), lheight = 0.7,id = c(F, F), margins = c(0,0.7,0.7))
bertin(especialista9, color=c("white", "darkblue"), xlim = c(0.3,0.4), ylim = c(.0, 0.5), lheight = 0.7,id = c(F, F), margins = c(0,0.7,0.7))
bertin(especialista10, color=c("white", "darkblue"), xlim = c(0.3,0.4), ylim = c(.0, 0.5), lheight = 0.7,id = c(F, F), margins = c(0,0.7,0.7))

# clusteriza construtos e os veículos para gerar dendogramas individuais
cluster(especialista1,lab.cex=1,xlim = c(40,0))
cluster(especialista2,lab.cex=1,xlim = c(40,0))
cluster(especialista3,lab.cex=1,xlim = c(40,0))
cluster(especialista4,lab.cex=1,xlim = c(40,0))
cluster(especialista5,lab.cex=1,xlim = c(40,0))
cluster(especialista6,lab.cex=1,xlim = c(40,0))
cluster(especialista7,lab.cex=1,xlim = c(40,0))
cluster(especialista8,lab.cex=1,xlim = c(40,0))
cluster(especialista9,lab.cex=1,xlim = c(40,0))
cluster(especialista10,lab.cex=1,xlim = c(40,0))

# gera gráficos Biplot simples individuais
biplotSimple(especialista1, c.label.col="darkblue", zoom = 0.9)
biplotSimple(especialista2, c.label.col="darkblue", zoom = 0.9)
biplotSimple(especialista3, c.label.col="darkblue", zoom = 0.9)
biplotSimple(especialista4, c.label.col="darkblue", zoom = 0.9)
biplotSimple(especialista5, c.label.col="darkblue", zoom = 0.9)
biplotSimple(especialista6, c.label.col="darkblue", zoom = 0.9)
biplotSimple(especialista7, c.label.col="darkblue", zoom = 0.9)
biplotSimple(especialista8, c.label.col="darkblue", zoom = 0.9)
biplotSimple(especialista9, c.label.col="darkblue", zoom = 0.9)
biplotSimple(especialista10, c.label.col="darkblue", zoom = 0.9)

# passo a passo para criar as visualizações individuais para os não-especialistas

# extrai subconjuntos do RGT
nao_especialista1<-rgt_grid_nesp[1:10, 1:5] 
nao_especialista2<-rgt_grid_nesp[11:20, 1:5]
nao_especialista3<-rgt_grid_nesp[21:30, 1:5]
nao_especialista4<-rgt_grid_nesp[31:40, 1:5]
nao_especialista5<-rgt_grid_nesp[41:50, 1:5]
nao_especialista6<-rgt_grid_nesp[51:60, 1:5]
nao_especialista7<-rgt_grid_nesp[61:70, 1:5]
nao_especialista8<-rgt_grid_nesp[71:80, 1:5]
nao_especialista9<-rgt_grid_nesp[81:90, 1:5]
nao_especialista10<-rgt_grid_nesp[91:100, 1:5]

# cria as matrizes de correlação individuais
matriz_correlacao_nesp1 <- constructCor(nao_especialista1)
matriz_correlacao_nesp2 <- constructCor(nao_especialista2)
matriz_correlacao_nesp3 <- constructCor(nao_especialista3) 
matriz_correlacao_nesp4 <- constructCor(nao_especialista4) 
matriz_correlacao_nesp5 <- constructCor(nao_especialista5) 
matriz_correlacao_nesp6 <- constructCor(nao_especialista6) 
matriz_correlacao_nesp7 <- constructCor(nao_especialista7) 
matriz_correlacao_nesp8 <- constructCor(nao_especialista8) 
matriz_correlacao_nesp9 <- constructCor(nao_especialista9) 
matriz_correlacao_nesp10 <- constructCor(nao_especialista10) 

# plota as matrizes de correlação individuais
corrplot(matriz_correlacao_nesp1, method ='number', type = "lower", number.cex = 0.8, tl.pos = "n", cl.cex = 0.8)  
corrplot(matriz_correlacao_nesp2, method ='number', type = "lower", number.cex = 0.8, tl.pos = "n", cl.cex = 0.8)
corrplot(matriz_correlacao_nesp3, method ='number', type = "lower", number.cex = 0.8, tl.pos = "n", cl.cex = 0.8)
corrplot(matriz_correlacao_nesp4, method ='number', type = "lower", number.cex = 0.8, tl.pos = "n", cl.cex = 0.8)
corrplot(matriz_correlacao_nesp5, method ='number', type = "lower", number.cex = 0.8, tl.pos = "n", cl.cex = 0.8)
corrplot(matriz_correlacao_nesp6, method ='number', type = "lower", number.cex = 0.8, tl.pos = "n", cl.cex = 0.8)
corrplot(matriz_correlacao_nesp7, method ='number', type = "lower", number.cex = 0.8, tl.pos = "n", cl.cex = 0.8)
corrplot(matriz_correlacao_nesp8, method ='number', type = "lower", number.cex = 0.8, tl.pos = "n", cl.cex = 0.8)
corrplot(matriz_correlacao_nesp9, method ='number', type = "lower", number.cex = 0.8, tl.pos = "n", cl.cex = 0.8)
corrplot(matriz_correlacao_nesp10, method ='number', type = "lower", number.cex = 0.8, tl.pos = "n", cl.cex = 0.8)

# gera matrizes ordenáveis de Bertin individuais
bertin(nao_especialista1, color=c("white", "darkred"), xlim = c(0.3,0.4), ylim = c(.0, 0.5), lheight = 0.7,id = c(F, F), margins = c(0,0.7,0.7)) 
bertin(nao_especialista2, color=c("white", "darkred"), xlim = c(0.3,0.4), ylim = c(.0, 0.5), lheight = 0.7,id = c(F, F), margins = c(0,0.7,0.7))
bertin(nao_especialista3, color=c("white", "darkred"), xlim = c(0.3,0.4), ylim = c(.0, 0.5), lheight = 0.7,id = c(F, F), margins = c(0,0.7,0.7))
bertin(nao_especialista4, color=c("white", "darkred"), xlim = c(0.3,0.4), ylim = c(.0, 0.5), lheight = 0.7,id = c(F, F), margins = c(0,0.7,0.7))
bertin(nao_especialista5, color=c("white", "darkred"), xlim = c(0.3,0.4), ylim = c(.0, 0.5), lheight = 0.7,id = c(F, F), margins = c(0,0.7,0.7))
bertin(nao_especialista6, color=c("white", "darkred"), xlim = c(0.3,0.4), ylim = c(.0, 0.5), lheight = 0.7,id = c(F, F), margins = c(0,0.7,0.7))
bertin(nao_especialista7, color=c("white", "darkred"), xlim = c(0.3,0.4), ylim = c(.0, 0.5), lheight = 0.7,id = c(F, F), margins = c(0,0.7,0.7))
bertin(nao_especialista8, color=c("white", "darkred"), xlim = c(0.3,0.4), ylim = c(.0, 0.5), lheight = 0.7,id = c(F, F), margins = c(0,0.7,0.7))
bertin(nao_especialista9, color=c("white", "darkred"), xlim = c(0.3,0.4), ylim = c(.0, 0.5), lheight = 0.7,id = c(F, F), margins = c(0,0.7,0.7))
bertin(nao_especialista10, color=c("white", "darkred"), xlim = c(0.3,0.4), ylim = c(.0, 0.5), lheight = 0.7,id = c(F, F), margins = c(0,0.7,0.7))

# clusteriza construtos e os veículos para gerar dendogramas individuais
cluster(nao_especialista1,lab.cex=1,xlim = c(40,0))
cluster(nao_especialista2,lab.cex=1,xlim = c(40,0))
cluster(nao_especialista3,lab.cex=1,xlim = c(40,0))
cluster(nao_especialista4,lab.cex=1,xlim = c(40,0))
cluster(nao_especialista5,lab.cex=1,xlim = c(40,0))
cluster(nao_especialista6,lab.cex=1,xlim = c(40,0))
cluster(nao_especialista7,lab.cex=1,xlim = c(40,0))
cluster(nao_especialista8,lab.cex=1,xlim = c(40,0))
cluster(nao_especialista9,lab.cex=1,xlim = c(40,0))
cluster(nao_especialista10,lab.cex=1,xlim = c(40,0))

# gera gráficos Biplot simples individuais
biplotSimple(nao_especialista1, c.label.col="darkred", zoom = 0.9)
biplotSimple(nao_especialista2, c.label.col="darkred", zoom = 0.9)
biplotSimple(nao_especialista3, c.label.col="darkred", zoom = 0.9)
biplotSimple(nao_especialista4, c.label.col="darkred", zoom = 0.9)
biplotSimple(nao_especialista5, c.label.col="darkred", zoom = 0.9)
biplotSimple(nao_especialista6, c.label.col="darkred", zoom = 0.9)
biplotSimple(nao_especialista7, c.label.col="darkred", zoom = 0.9)
biplotSimple(nao_especialista8, c.label.col="darkred", zoom = 0.9)
biplotSimple(nao_especialista9, c.label.col="darkred", zoom = 0.9)
biplotSimple(nao_especialista10, c.label.col="darkred", zoom = 0.9)