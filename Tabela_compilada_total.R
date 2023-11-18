# Carregando as bibliotecas necess�rias
library(tidyverse)
library(readxl)
library(writexl)

# Fun��o para abrir arquivos relacionados � AV_Bimestral_2
abrir_av_bimestral_2 <- function(caminho, filtro) {
  # Listando todos os arquivos no caminho especificado
  arquivos <- list.files(caminho, full.names = TRUE, recursive = TRUE)
  # Filtrando os arquivos que correspondem ao filtro
  arquivos <- arquivos[!grepl(filtro, arquivos)]
  # Selecionando pastas que correspondem a padr�es espec�ficos nos nomes dos arquivos
  pastas_av_bimestral_2 <- grep("AV_Bimestral_2|AV_2_Bimestral|AV_Bimesrtal_2|Pactuacao_AV_2", arquivos, value = TRUE)
  # Criando uma tibble com os nomes das pastas encontradas
  arquivos <- tibble(arquivo = pastas_av_bimestral_2)
  
  return(arquivos)
}

# Fun��o para abrir arquivos relacionados � AV_Bimestral_1
abrir_av_bimestral_1 <- function(caminho, filtro) {
  # Listando todos os arquivos no caminho especificado
  arquivos <- list.files(caminho, full.names = TRUE, recursive = TRUE)
  # Filtrando os arquivos que correspondem ao filtro
  arquivos <- arquivos[!grepl(filtro, arquivos)]
  # Selecionando arquivos com extens�o ".xlsm"
  pastas_av1 <- arquivos[grepl(".xlsm", arquivos)]
  # Criando uma tibble com os nomes dos arquivos encontrados
  arquivos <- tibble(arquivo = pastas_av1)
  
  return(arquivos)
}

# Fun��o para ler uma planilha de dados espec�fica
ler_planilha_dados <- function(caminho_completo) {
  # Lendo a planilha Excel especificada e extraindo um intervalo de c�lulas
  cabe�alho <- read_excel(caminho_completo, sheet = "Principal", range = "b8:c19")
  # Realizando algumas manipula��es nos dados lidos
  cabe�alho <- cabe�alho %>%
    t() %>%
    as_tibble() %>%
    rownames_to_column("value") %>%
    `colnames<-`(.[1,]) %>%
    .[-1,] %>%
    `rownames<-`(NULL)
  
  return(cabe�alho)
}

# Fun��o para ler notas de servidores a partir de uma planilha
notas_servidores <- function(caminho_completo){
  # Lendo a planilha Excel especificada e extraindo um intervalo de c�lulas
  nota <- read_excel(caminho_completo, sheet = "Avalia��o", range = "C44:E62") 
  # Realizando algumas manipula��es nos dados lidos
  nota <- nota %>% t() %>% 
    as_tibble() %>% 
    rownames_to_column("value") %>% 
    `colnames<-`(.[1,]) %>% 
    .[-c(1, 2),] %>%
    `rownames<-`(NULL) 
  
  # Chamando a fun��o ler_planilha_dados para obter um cabe�alho
  cabe�alho <- ler_planilha_dados(caminho_completo)
  # Chamando a fun��o qualitativo (que n�o est� definida no c�digo) para obter escala
  escala <- qualitativo(caminho_completo)
  # Combinando os dados do cabe�alho, nota e escala em uma �nica tibble
  nota <- bind_cols(cabe�alho, nota, escala)
  
  return(nota)
}

# Fun��o para processar os resultados da AV_Bimestral_2
resultado_Av2 <- function(av2)
{
  # Lendo a planilha Excel especificada com tipos de coluna especificados
  avalia��o_av2 <- read_excel(av2, sheet = "ResultadoDaAvParcial", col_types = c("numeric", "numeric", "date", "text", "text",
                                                                                 "text", "text", "text", "text", "text", "numeric",
                                                                                 "numeric", "numeric","text", "text"))
  # Chamando a fun��o Avalia��o_Qualitativa para fazer mais manipula��es nos dados
  avalia��o_av2 <- Avalia��o_Qualitativa(av2, avalia��o_av2)
  
  return(avalia��o_av2)
}

# Fun��o para processar os resultados da AV_Bimestral_1
resultado_Av1 <- function(av1)
{
  # Lendo a planilha Excel especificada com tipos de coluna especificados
  avalia��o_av1 <- read_excel(av1, sheet = "TabelaCompilada", col_types = c("numeric", "numeric", "date", "text", "text",
                                                                            "text", "text", "text", "text","numeric","numeric", 
                                                                            "numeric", "text"))
  # Chamando a fun��o Avalia��o_Qualitativa para fazer mais manipula��es nos dados
  avalia��o_av1 <- Avalia��o_Qualitativa(av1, avalia��o_av1)
  
  return(avalia��o_av1)
}

# Fun��o para processar dados qualitativos de avalia��o
Avalia��o_Qualitativa <- function(caminho, cabe�alho)
{
  # Lendo a planilha Excel especificada e extraindo um intervalo de c�lulas
  Qualitativa <- t(read_excel(caminho, sheet = "Avalia��o", range = "C44:E63")) %>% as_tibble
  colnames(Qualitativa) <- Qualitativa[1,]
  Qualitativa <- Qualitativa[-c(1,2),]
  # Combinando os dados lidos com o cabe�alho fornecido
  Qualitativa <- bind_cols(cabe�alho, Qualitativa)
  
  return(Qualitativa)
}

# Fun��o para realizar a limpeza dos dados em uma tabela compilada
limpeza_dos_dados <- function(tabela_compilada)
{
  # Realizando uma jun��o completa entre tabelas tabela_av1 e tabela_av2 (n�o definidas no c�digo)
  tabela_compilada$`DATA INICIO ESTAGIO PROB` <- as.numeric(tabela_compilada$`DATA INICIO ESTAGIO PROB`)
  tabela_compilada$`DATA INICIO ESTAGIO PROB` <- as.Date(tabela_compilada$`DATA INICIO ESTAGIO PROB`, origin = "1899-12-30")
  # Filtrando dados onde RESULTADO n�o � NA e maior que zero
  tabela_compilada<- tabela_compilada %>% filter(!is.na(RESULTADO) & RESULTADO > 0)
  # Substituindo "02/" por nada na coluna MATRICULA
  tabela_compilada$MATRICULA <- str_replace_all(tabela_compilada$MATRICULA, "02/","")
  
  # Realizando uma transforma��o na coluna `[Servidor] COMUNICA��O`
  tabela_compilada<- tabela_compilada %>%
    mutate(`[Servidor] COMUNICA��O` = ifelse(!is.na(`[Servidor] FALAR BEM EM P�BLICO`), `[Servidor] FALAR BEM EM P�BLICO`, `[Servidor] COMUNICA��O`)) %>% select(-16)
  
  # Selecionando colunas espec�ficas
  tabela_compilada <- select(tabela_compilada, -c(32,34:61))
  
  # Reordenando as colunas da tabela
  tabela_compilada <- tabela_compilada %>% select(1:14, 32, everything())
  
  # Realizando uma transforma��o em algumas colunas espec�ficas
  tabela_compilada <- tabela_compilada %>%
    mutate(across(14:25, ~ifelse(PERFIL == "T" & !is.na(`[Servidor] QUALIDADE`), NA, .)))
  
  return(tabela_compilada)
}

# Fun��o para criar um arquivo Excel a partir de uma tabela compilada
criar_excel <- function(tabela_compilada)
{
  write_xlsx(tabela_compilada, "F:/CDP/Avalia��o de Desempenho/Ciclo 2023/tabela_compilada_total.xlsx") 
}

# Fun��o para multiplicar valores em colunas espec�ficas
multiplicar <- function(colunas)
{
  colunas <- 2 * colunas
  return (colunas)
}

# Fun��o para processar indicadores
indicadores <- function(cesta_de_indicadores)
{
  # adicionando ao ANO a vari�vel N�AVALIA��O e selecionando colunas espec�ficas
    cesta_de_indicadores <- cesta_de_indicadores %>%
    mutate(ANO = paste(ANO, N�AVALIA��O, sep = ".")) %>%
    select(-1) %>%
    # Convertendo colunas 9 a 27 para num�ricas
    mutate(across(9:27, as.numeric)) %>%
    # Multiplicando valores em colunas espec�ficas se o ANO for "2022.1" ou "2022.2"
    mutate(across(9:27, ~ifelse(ANO == "2022.1" | ANO == "2022.2", multiplicar(.), .)))  %>%
      filter(ANO != "2023.2") %>% mutate(RESULTADO = RESULTADO * 100)
    
    return(cesta_de_indicadores)
}

# Fun��o para criar um gr�fico de histograma
valores_grafico_histograma <- function(cesta_de_indicadores)
{
  # Filtrar os dados para o per�odo "2023.1" e calcular a m�dia por SETOR
  top_bottom_sem_100 <- cesta_de_indicadores %>%
    filter(ANO == "2023.1") %>%
    group_by(SETOR) %>%
    summarise(media = mean(RESULTADO)) %>%
    arrange(desc(media)) %>% filter(media != 100)
  
  # Definir a ordem dos fatores no gr�fico com base nas m�dias
  top_bottom_sem_100$SETOR <- factor(top_bottom_sem_100$SETOR, levels = unique(top_bottom_sem_100$SETOR))
  
  # Selecionar os cinco melhores e os cinco piores SETORES
  top_bottom_sem_100 <- top_bottom_sem_100 %>%
    slice(c(1:5, (n() - 4):n()))
  
  # Criar um grupo com base na m�dia em rela��o � m�dia geral
  top_bottom_sem_100$Grupo <- ifelse(top_bottom_sem_100$media >= mean(cesta_de_indicadores$RESULTADO), "Os cinco melhores", "Os cinco piores")
  
  # Chamar a fun��o grafico_histograma e passar os dados processados
  grafico_histograma(top_bottom_sem_100)
}

# Fun��o para criar um gr�fico de histograma com base em dados processados
grafico_histograma <- function(dados_media)
{
  # Definir cores para os grupos
  cores <- c("Os cinco melhores" = "green", "Os cinco piores" = "red")
  
  # Criar o gr�fico de barras
  media_top5_e_bottom5 <- ggplot(dados_media, aes(x = reorder(SETOR, media), y = media, fill = Grupo)) +
    geom_bar(stat = "identity") +
    scale_fill_manual(values = cores) + 
    geom_hline(yintercept = mean(cesta_de_indicadores$RESULTADO), color = "red", linetype = "dashed") + 
    labs(title = "Os cinco setores com as melhores e piores m�dias em 2023.1",
         x = "Setor",
         y = "M�dia de Resultado (%)") +
    geom_text(aes(label = sprintf("%.2f", media)), vjust = -0.5, size = 4) + 
    annotate("text", x = 1, y = mean(cesta_de_indicadores$RESULTADO), label = sprintf("%.3f", mean(cesta_de_indicadores$RESULTADO)), vjust = -0.5, color = "red") +
    coord_cartesian(ylim = c(75, 100))
  
  return(media_top5_e_bottom5)
}

# Fun��o para criar um gr�fico de boxplot
grafico_boxplot <- function(cesta_de_indicadores)
{
  # Criar o gr�fico de boxplot para visualizar a distribui��o de resultados por per�odos avaliativos
  distribuicao_resultado <- ggplot(cesta_de_indicadores, aes(x = ANO, y = RESULTADO)) + 
    geom_boxplot(fill = "dodgerblue", color = "black", alpha = 1) +
    labs(title = "Distribui��o de Resultados por Per�odos Avaliativos", 
         x = "Per�odo Avaliativo",
         y = "Resultado (%)") +
    theme_minimal() +
    theme(axis.title.x = element_text(size = 12),
          axis.title.y = element_text(size = 12),
          axis.text.x = element_text(size = 10, angle = 45, hjust = 1),
          axis.text.y = element_text(size = 10),
          plot.title = element_text(size = 14, hjust = 0.5)) +
    scale_fill_manual(values = c("dodgerblue")) +
    geom_hline(yintercept = mean(cesta_de_indicadores$RESULTADO), 
               linetype = "dashed", color = "red", size = 1) +
    theme(legend.position = "none") +
    coord_cartesian(ylim = c(70, 100))
  
  return(distribuicao_resultado)
}

# Fun��o para criar um gr�fico de linha (s�rie temporal)
grafico_linha <- function(cesta_de_indicadores)
{
  # Calcular m�dia e desvio padr�o por ANO
  serie_temporal <- cesta_de_indicadores %>% group_by(ANO) %>% summarise(media = mean(RESULTADO),
                                                                         desvio_padrao = sd(RESULTADO))
  
  # Criar o gr�fico de linha (s�rie temporal) com m�dia e desvio padr�o
  media_e_desvio_padrao_pelo_tempo <- ggplot(serie_temporal, aes(x = ANO, y = media, group = 1)) +
    geom_point(color = "blue", size = 3) +
    geom_line(color = "red") +
    geom_text(aes(label = sprintf("%.3f", desvio_padrao)), vjust = -1, hjust = 0.5, size = 4, color = "black") +
    geom_text(aes(label = sprintf("%.1f", media)), vjust = 2, hjust = 0.4, size = 5, color = "black") +
    labs(x = "Ano", y = "M�dia de Valores", title = "S�rie Temporal e Desvio Padr�o") +
    theme_classic() + 
    coord_cartesian(ylim = c(95,97))
  
  return(media_e_desvio_padrao_pelo_tempo)
}

# Definir caminhos e filtros para buscar os arquivos
caminho <- c("M:/","F:/CDP/Avalia��o de Desempenho/Ciclo 2022/BACKUP")
filtro_av1 <- c("Banco|Banco2|Servidor|~\\$|.pdf|_MATRIZ PRODUTOS|MODELO|TabelaCompilada|AV_Bimestral_2|Avalia��o 1� ciclo|Pactua��o 2� Ciclo|202209-202208|Matriz|AV_2_Bimestral|GC6preenchida|AV_Bimesrtal_2|STE-30-05-2023|M://GAP/2022/|Pactuacao_AV_2|PARA AVALIA��O SUBJETIVA")
filtro_av2 <- c("~\\$|Respons�veis de �rea|historico|pdf|PDF|_MATRIZ PRODUTOS|MODELO|TabelaCompilada|GC6preenchida|Tabela_Compilada|Matriz_produtos|Tabela Compilada|.txt|M://SUBPES/AV_Bimestral_2/SUB-SEGURIDADE/")

# Abrir pastas relacionadas � AV_Bimestral_1 e AV_Bimestral_2
pastas_av1 <- abrir_av_bimestral_1(caminho, filtro_av1)
pastas_av2 <- abrir_av_bimestral_2(caminho, filtro_av2)

# Ler e processar os dados dos arquivos da AV_Bimestral_1 e AV_Bimestral_2
tabela_av1 <- map(pastas_av1$arquivo, resultado_Av1) %>% bind_rows()
tabela_av2 <- map(pastas_av2$arquivo, resultado_Av2) %>% bind_rows()

# Combinar as tabelas de AV_Bimestral_1 e AV_Bimestral_2 e realizar limpeza dos dados
tabela_compilada <- full_join(tabela_av1, tabela_av2) %>% limpeza_dos_dados() 

# Calcular indicadores a partir dos dados da tabela compilada
cesta_de_indicadores <- indicadores(select(tabela_compilada, -c(3,4,9,13)))

# Criar um arquivo Excel com os dados processados
criar_excel(tabela_compilada)

# Gerar gr�fico de histograma com os cinco melhores e cinco piores setores
histograma_top5_e_bottom5 <- valores_grafico_histograma(cesta_de_indicadores)

# Gerar gr�fico de boxplot para visualizar a distribui��o de resultados por per�odos avaliativos
metricas_dos_resultados <- grafico_boxplot(cesta_de_indicadores)

# Gerar gr�fico de linha (s�rie temporal) com m�dias e desvio padr�o
serie_temporal_resultado_total <- grafico_linha(cesta_de_indicadores)

# Exibir os gr�ficos
plot(histograma_top5_e_bottom5)
plot(metricas_dos_resultados)
plot(serie_temporal_resultado_total)

