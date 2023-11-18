# Carregando as bibliotecas necessárias
library(tidyverse)
library(readxl)
library(writexl)

# Função para abrir arquivos relacionados à AV_Bimestral_2
abrir_av_bimestral_2 <- function(caminho, filtro) {
  # Listando todos os arquivos no caminho especificado
  arquivos <- list.files(caminho, full.names = TRUE, recursive = TRUE)
  # Filtrando os arquivos que correspondem ao filtro
  arquivos <- arquivos[!grepl(filtro, arquivos)]
  # Selecionando pastas que correspondem a padrões específicos nos nomes dos arquivos
  pastas_av_bimestral_2 <- grep("AV_Bimestral_2|AV_2_Bimestral|AV_Bimesrtal_2|Pactuacao_AV_2", arquivos, value = TRUE)
  # Criando uma tibble com os nomes das pastas encontradas
  arquivos <- tibble(arquivo = pastas_av_bimestral_2)
  
  return(arquivos)
}

# Função para abrir arquivos relacionados à AV_Bimestral_1
abrir_av_bimestral_1 <- function(caminho, filtro) {
  # Listando todos os arquivos no caminho especificado
  arquivos <- list.files(caminho, full.names = TRUE, recursive = TRUE)
  # Filtrando os arquivos que correspondem ao filtro
  arquivos <- arquivos[!grepl(filtro, arquivos)]
  # Selecionando arquivos com extensão ".xlsm"
  pastas_av1 <- arquivos[grepl(".xlsm", arquivos)]
  # Criando uma tibble com os nomes dos arquivos encontrados
  arquivos <- tibble(arquivo = pastas_av1)
  
  return(arquivos)
}

# Função para ler uma planilha de dados específica
ler_planilha_dados <- function(caminho_completo) {
  # Lendo a planilha Excel especificada e extraindo um intervalo de células
  cabeçalho <- read_excel(caminho_completo, sheet = "Principal", range = "b8:c19")
  # Realizando algumas manipulações nos dados lidos
  cabeçalho <- cabeçalho %>%
    t() %>%
    as_tibble() %>%
    rownames_to_column("value") %>%
    `colnames<-`(.[1,]) %>%
    .[-1,] %>%
    `rownames<-`(NULL)
  
  return(cabeçalho)
}

# Função para ler notas de servidores a partir de uma planilha
notas_servidores <- function(caminho_completo){
  # Lendo a planilha Excel especificada e extraindo um intervalo de células
  nota <- read_excel(caminho_completo, sheet = "Avaliação", range = "C44:E62") 
  # Realizando algumas manipulações nos dados lidos
  nota <- nota %>% t() %>% 
    as_tibble() %>% 
    rownames_to_column("value") %>% 
    `colnames<-`(.[1,]) %>% 
    .[-c(1, 2),] %>%
    `rownames<-`(NULL) 
  
  # Chamando a função ler_planilha_dados para obter um cabeçalho
  cabeçalho <- ler_planilha_dados(caminho_completo)
  # Chamando a função qualitativo (que não está definida no código) para obter escala
  escala <- qualitativo(caminho_completo)
  # Combinando os dados do cabeçalho, nota e escala em uma única tibble
  nota <- bind_cols(cabeçalho, nota, escala)
  
  return(nota)
}

# Função para processar os resultados da AV_Bimestral_2
resultado_Av2 <- function(av2)
{
  # Lendo a planilha Excel especificada com tipos de coluna especificados
  avaliação_av2 <- read_excel(av2, sheet = "ResultadoDaAvParcial", col_types = c("numeric", "numeric", "date", "text", "text",
                                                                                 "text", "text", "text", "text", "text", "numeric",
                                                                                 "numeric", "numeric","text", "text"))
  # Chamando a função Avaliação_Qualitativa para fazer mais manipulações nos dados
  avaliação_av2 <- Avaliação_Qualitativa(av2, avaliação_av2)
  
  return(avaliação_av2)
}

# Função para processar os resultados da AV_Bimestral_1
resultado_Av1 <- function(av1)
{
  # Lendo a planilha Excel especificada com tipos de coluna especificados
  avaliação_av1 <- read_excel(av1, sheet = "TabelaCompilada", col_types = c("numeric", "numeric", "date", "text", "text",
                                                                            "text", "text", "text", "text","numeric","numeric", 
                                                                            "numeric", "text"))
  # Chamando a função Avaliação_Qualitativa para fazer mais manipulações nos dados
  avaliação_av1 <- Avaliação_Qualitativa(av1, avaliação_av1)
  
  return(avaliação_av1)
}

# Função para processar dados qualitativos de avaliação
Avaliação_Qualitativa <- function(caminho, cabeçalho)
{
  # Lendo a planilha Excel especificada e extraindo um intervalo de células
  Qualitativa <- t(read_excel(caminho, sheet = "Avaliação", range = "C44:E63")) %>% as_tibble
  colnames(Qualitativa) <- Qualitativa[1,]
  Qualitativa <- Qualitativa[-c(1,2),]
  # Combinando os dados lidos com o cabeçalho fornecido
  Qualitativa <- bind_cols(cabeçalho, Qualitativa)
  
  return(Qualitativa)
}

# Função para realizar a limpeza dos dados em uma tabela compilada
limpeza_dos_dados <- function(tabela_compilada)
{
  # Realizando uma junção completa entre tabelas tabela_av1 e tabela_av2 (não definidas no código)
  tabela_compilada$`DATA INICIO ESTAGIO PROB` <- as.numeric(tabela_compilada$`DATA INICIO ESTAGIO PROB`)
  tabela_compilada$`DATA INICIO ESTAGIO PROB` <- as.Date(tabela_compilada$`DATA INICIO ESTAGIO PROB`, origin = "1899-12-30")
  # Filtrando dados onde RESULTADO não é NA e maior que zero
  tabela_compilada<- tabela_compilada %>% filter(!is.na(RESULTADO) & RESULTADO > 0)
  # Substituindo "02/" por nada na coluna MATRICULA
  tabela_compilada$MATRICULA <- str_replace_all(tabela_compilada$MATRICULA, "02/","")
  
  # Realizando uma transformação na coluna `[Servidor] COMUNICAÇÃO`
  tabela_compilada<- tabela_compilada %>%
    mutate(`[Servidor] COMUNICAÇÃO` = ifelse(!is.na(`[Servidor] FALAR BEM EM PÚBLICO`), `[Servidor] FALAR BEM EM PÚBLICO`, `[Servidor] COMUNICAÇÃO`)) %>% select(-16)
  
  # Selecionando colunas específicas
  tabela_compilada <- select(tabela_compilada, -c(32,34:61))
  
  # Reordenando as colunas da tabela
  tabela_compilada <- tabela_compilada %>% select(1:14, 32, everything())
  
  # Realizando uma transformação em algumas colunas específicas
  tabela_compilada <- tabela_compilada %>%
    mutate(across(14:25, ~ifelse(PERFIL == "T" & !is.na(`[Servidor] QUALIDADE`), NA, .)))
  
  return(tabela_compilada)
}

# Função para criar um arquivo Excel a partir de uma tabela compilada
criar_excel <- function(tabela_compilada)
{
  write_xlsx(tabela_compilada, "F:/CDP/Avaliação de Desempenho/Ciclo 2023/tabela_compilada_total.xlsx") 
}

# Função para multiplicar valores em colunas específicas
multiplicar <- function(colunas)
{
  colunas <- 2 * colunas
  return (colunas)
}

# Função para processar indicadores
indicadores <- function(cesta_de_indicadores)
{
  # adicionando ao ANO a variável NºAVALIAÇÃO e selecionando colunas específicas
    cesta_de_indicadores <- cesta_de_indicadores %>%
    mutate(ANO = paste(ANO, NºAVALIAÇÃO, sep = ".")) %>%
    select(-1) %>%
    # Convertendo colunas 9 a 27 para numéricas
    mutate(across(9:27, as.numeric)) %>%
    # Multiplicando valores em colunas específicas se o ANO for "2022.1" ou "2022.2"
    mutate(across(9:27, ~ifelse(ANO == "2022.1" | ANO == "2022.2", multiplicar(.), .)))  %>%
      filter(ANO != "2023.2") %>% mutate(RESULTADO = RESULTADO * 100)
    
    return(cesta_de_indicadores)
}

# Função para criar um gráfico de histograma
valores_grafico_histograma <- function(cesta_de_indicadores)
{
  # Filtrar os dados para o período "2023.1" e calcular a média por SETOR
  top_bottom_sem_100 <- cesta_de_indicadores %>%
    filter(ANO == "2023.1") %>%
    group_by(SETOR) %>%
    summarise(media = mean(RESULTADO)) %>%
    arrange(desc(media)) %>% filter(media != 100)
  
  # Definir a ordem dos fatores no gráfico com base nas médias
  top_bottom_sem_100$SETOR <- factor(top_bottom_sem_100$SETOR, levels = unique(top_bottom_sem_100$SETOR))
  
  # Selecionar os cinco melhores e os cinco piores SETORES
  top_bottom_sem_100 <- top_bottom_sem_100 %>%
    slice(c(1:5, (n() - 4):n()))
  
  # Criar um grupo com base na média em relação à média geral
  top_bottom_sem_100$Grupo <- ifelse(top_bottom_sem_100$media >= mean(cesta_de_indicadores$RESULTADO), "Os cinco melhores", "Os cinco piores")
  
  # Chamar a função grafico_histograma e passar os dados processados
  grafico_histograma(top_bottom_sem_100)
}

# Função para criar um gráfico de histograma com base em dados processados
grafico_histograma <- function(dados_media)
{
  # Definir cores para os grupos
  cores <- c("Os cinco melhores" = "green", "Os cinco piores" = "red")
  
  # Criar o gráfico de barras
  media_top5_e_bottom5 <- ggplot(dados_media, aes(x = reorder(SETOR, media), y = media, fill = Grupo)) +
    geom_bar(stat = "identity") +
    scale_fill_manual(values = cores) + 
    geom_hline(yintercept = mean(cesta_de_indicadores$RESULTADO), color = "red", linetype = "dashed") + 
    labs(title = "Os cinco setores com as melhores e piores médias em 2023.1",
         x = "Setor",
         y = "Média de Resultado (%)") +
    geom_text(aes(label = sprintf("%.2f", media)), vjust = -0.5, size = 4) + 
    annotate("text", x = 1, y = mean(cesta_de_indicadores$RESULTADO), label = sprintf("%.3f", mean(cesta_de_indicadores$RESULTADO)), vjust = -0.5, color = "red") +
    coord_cartesian(ylim = c(75, 100))
  
  return(media_top5_e_bottom5)
}

# Função para criar um gráfico de boxplot
grafico_boxplot <- function(cesta_de_indicadores)
{
  # Criar o gráfico de boxplot para visualizar a distribuição de resultados por períodos avaliativos
  distribuicao_resultado <- ggplot(cesta_de_indicadores, aes(x = ANO, y = RESULTADO)) + 
    geom_boxplot(fill = "dodgerblue", color = "black", alpha = 1) +
    labs(title = "Distribuição de Resultados por Períodos Avaliativos", 
         x = "Período Avaliativo",
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

# Função para criar um gráfico de linha (série temporal)
grafico_linha <- function(cesta_de_indicadores)
{
  # Calcular média e desvio padrão por ANO
  serie_temporal <- cesta_de_indicadores %>% group_by(ANO) %>% summarise(media = mean(RESULTADO),
                                                                         desvio_padrao = sd(RESULTADO))
  
  # Criar o gráfico de linha (série temporal) com média e desvio padrão
  media_e_desvio_padrao_pelo_tempo <- ggplot(serie_temporal, aes(x = ANO, y = media, group = 1)) +
    geom_point(color = "blue", size = 3) +
    geom_line(color = "red") +
    geom_text(aes(label = sprintf("%.3f", desvio_padrao)), vjust = -1, hjust = 0.5, size = 4, color = "black") +
    geom_text(aes(label = sprintf("%.1f", media)), vjust = 2, hjust = 0.4, size = 5, color = "black") +
    labs(x = "Ano", y = "Média de Valores", title = "Série Temporal e Desvio Padrão") +
    theme_classic() + 
    coord_cartesian(ylim = c(95,97))
  
  return(media_e_desvio_padrao_pelo_tempo)
}

# Definir caminhos e filtros para buscar os arquivos
caminho <- c("M:/","F:/CDP/Avaliação de Desempenho/Ciclo 2022/BACKUP")
filtro_av1 <- c("Banco|Banco2|Servidor|~\\$|.pdf|_MATRIZ PRODUTOS|MODELO|TabelaCompilada|AV_Bimestral_2|Avaliação 1º ciclo|Pactuação 2º Ciclo|202209-202208|Matriz|AV_2_Bimestral|GC6preenchida|AV_Bimesrtal_2|STE-30-05-2023|M://GAP/2022/|Pactuacao_AV_2|PARA AVALIAÇÃO SUBJETIVA")
filtro_av2 <- c("~\\$|Responsáveis de Área|historico|pdf|PDF|_MATRIZ PRODUTOS|MODELO|TabelaCompilada|GC6preenchida|Tabela_Compilada|Matriz_produtos|Tabela Compilada|.txt|M://SUBPES/AV_Bimestral_2/SUB-SEGURIDADE/")

# Abrir pastas relacionadas à AV_Bimestral_1 e AV_Bimestral_2
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

# Gerar gráfico de histograma com os cinco melhores e cinco piores setores
histograma_top5_e_bottom5 <- valores_grafico_histograma(cesta_de_indicadores)

# Gerar gráfico de boxplot para visualizar a distribuição de resultados por períodos avaliativos
metricas_dos_resultados <- grafico_boxplot(cesta_de_indicadores)

# Gerar gráfico de linha (série temporal) com médias e desvio padrão
serie_temporal_resultado_total <- grafico_linha(cesta_de_indicadores)

# Exibir os gráficos
plot(histograma_top5_e_bottom5)
plot(metricas_dos_resultados)
plot(serie_temporal_resultado_total)


