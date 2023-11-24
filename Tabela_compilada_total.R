  library(tidyverse)
  library(readxl)
  library(writexl)
  
  # Função para abrir arquivos relacionados á AV_Bimestral_2
  abrir_av_bimestral_2 <- function(caminho, filtro) {
    # Listando todos os arquivos no caminho especificado
    arquivos <- list.files(caminho, full.names = TRUE, recursive = TRUE)
    # Filtrando os arquivos que correspondem ao filtro
    arquivos <- arquivos[!grepl(filtro, arquivos)]
    # Selecionando pastas que correspondem a padrÃµes específicos nos nomes dos arquivos
    pastas_av_bimestral_2 <- grep("AV_Bimestral_2|AV_2_Bimestral|AV_Bimesrtal_2|Pactuacao_AV_2", arquivos, value = TRUE)
    # Criando uma tibble com os nomes das pastas encontradas
    arquivos <- tibble(arquivo = pastas_av_bimestral_2)
    
    return(arquivos)
  }
  
  # Função para abrir arquivos relacionados á AV_Bimestral_1
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
    cabecalho <- read_excel(caminho_completo, sheet = "Principal", range = "b8:c19")
    # Realizando algumas manipulações nos dados lidos
    cabecalho <- cabecalho %>%
      t() %>%
      as_tibble() %>%
      rownames_to_column("value") %>%
      `colnames<-`(.[1,]) %>%
      .[-1,] %>%
      `rownames<-`(NULL)
    
    return(cabecalho)
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
    
    # Chamando a Função ler_planilha_dados para obter um cabeÃ§alho
    cabecalho <- ler_planilha_dados(caminho_completo)
    # Chamando a Função qualitativo (que não está definida no código) para obter a escala
    escala <- qualitativo(caminho_completo)
    # Combinando os dados do cabeÃ§alho, nota e escala em uma única tibble
    nota <- bind_cols(cabecalho, nota, escala)
    
    return(nota)
  }
  
  # Função para processar os resultados da AV_Bimestral_2
  resultado_Av2 <- function(av2)
  {
    # Lendo a planilha Excel especificada com tipos de coluna especificados
    avaliacao_av2 <- read_excel(av2, sheet = "ResultadoDaAvParcial", col_types = c("numeric", "numeric", "date", "text", "text",
                                                                                   "text", "text", "text", "text", "text", "numeric",
                                                                                   "numeric", "numeric","text", "text"))
    # Chamando a Função Avaliação_Qualitativa para fazer mais manipulações nos dados
    avaliacao_av2 <- Avaliacao_Qualitativa(av2, avaliacao_av2)
    
    return(avaliacao_av2)
  }
  
  # Função para processar os resultados da AV_Bimestral_1
  resultado_Av1 <- function(av1)
  {
    # Lendo a planilha Excel especificada com tipos de coluna especificados
    avaliacao_av1 <- read_excel(av1, sheet = "TabelaCompilada", col_types = c("numeric", "numeric", "date", "text", "text",
                                                                              "text", "text", "text", "text", "numeric",
                                                                              "numeric", "numeric", "text"))
    # Chamando a Função Avaliação_Qualitativa para fazer mais manipulações nos dados
    avaliacao_av1 <- Avaliacao_Qualitativa(av1, avaliacao_av1)
    
    return(avaliacao_av1)
  }
  
  # Função para processar dados qualitativos de Avaliação
  Avaliacao_Qualitativa <- function(caminho, cabecalho)
  {
    # Lendo a planilha Excel especificada e extraindo um intervalo de células
    Qualitativa <- t(read_excel(caminho, sheet = "Avaliação", range = "C44:E63")) %>% as_tibble
    colnames(Qualitativa) <- Qualitativa[1,]
    Qualitativa <- Qualitativa[-c(1,2),]
    # Combinando os dados lidos com o cabeçalho fornecido
    Qualitativa <- bind_cols(cabecalho, Qualitativa)
    
    return(Qualitativa)
  }
  
  # Função para processar dados qualitativos de Avaliação (adicionada posteriormente)
  qualitativo <- function(caminho)
  {
    # Lendo a planilha Excel especificada e extraindo um intervalo de células
    Qualitativo <- read_excel(caminho, sheet = "Avaliação", range = "C44:E63")
    # Realizando algumas manipulações nos dados lidos
    Qualitativo <- Qualitativo %>%
      t() %>%
      as_tibble() %>%
      rownames_to_column("value") %>%
      `colnames<-`(.[1,]) %>%
      .[-c(1, 2),] %>%
      `rownames<-`(NULL)
    
    return(Qualitativo)
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
        
        # Realizando uma transformação na coluna `[Servidor] COMUNICAÇÃO
        tabela_compilada<- tabela_compilada %>%
          mutate(`[Servidor] COMUNICAÇÃO` = ifelse(!is.na(`[Servidor] FALAR BEM EM PÚBLICO`), `[Servidor] FALAR BEM EM PÚBLICO`, `[Servidor] COMUNICAÇÃO`)) %>% select(-16)
        
         #Selecionando colunas específicas
        tabela_compilada <- select(tabela_compilada, -c(32,34:42))
        
         #Reordenando as colunas da tabela
        tabela_compilada <- tabela_compilada %>% select(1:14, 32, everything())
        
        # Realizando uma transformação em algumas colunas específicas
        tabela_compilada <- tabela_compilada %>% mutate(ANO = paste(ANO, NºAVALIAÇÃO, sep = ".")) %>%
          select(-1) %>%
          mutate(across(13:21, ~ifelse(PERFIL == "E" & any(!is.na(.)), NA, .)),
                 across(22:30, ~ifelse(PERFIL == "S" & any(!is.na(.)), NA, .)),
                 across(26:31, ~ifelse(PERFIL == "E" & any(!is.na(.)), NA, .))) %>% distinct() %>%
          mutate(across(13:31, as.numeric)) %>% mutate(across(13:31, ~ifelse(ANO == "2022.1"| ANO == "2022.2", multiplicar(.), .))) %>%  
          mutate(SETOR = case_when(SETOR == "CAP1" ~ "1ªCAP",
                                   SETOR == "CAP2" ~ "2ªCAP",
                                   SETOR == "CAP3" ~ "3ªCAP",
                                   SETOR == "ECG/TCE-RJ" ~ "ECG",
                                   TRUE ~ SETOR)) 
        
        return(tabela_compilada)
      }
      transformacao_dados <- function(tabela_compilada)
      {
      tabela_compilada <- tabela_compilada %>% filter((SETOR == "GC5" | SETOR == "GP4" | SETOR == "GP5") & ANO == "2023.2") %>%
        mutate(ANO = "2023.1") %>%
        bind_rows(filter(tabela_compilada, ANO != "2023.2"))
      
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
      
      # Função para criar um gráfico de histograma
      valores_grafico_histograma <- function(tabela_compilada)
      {
        # Filtrar os dados para o período "2023.1" e calcular a média por SETOR
        top_bottom_sem_100 <- tabela_compilada %>%
          filter(ANO == "2023.1") %>%
          group_by(SETOR) %>%
          summarise(media = 100 * mean(RESULTADO)) %>%
          arrange(desc(media)) %>% filter(media != 100)
        
        # Definir a ordem dos fatores no gráfico com base nas médias
        top_bottom_sem_100$SETOR <- factor(top_bottom_sem_100$SETOR, levels = unique(top_bottom_sem_100$SETOR))
        
        # Selecionar os cinco melhores e os cinco piores SETORES
        top_bottom_sem_100 <- top_bottom_sem_100 %>%
          slice(c(1:5, (n() - 4):n()))
        
        # Criar um grupo com base na média em relação à média geral
        top_bottom_sem_100$Grupo <- ifelse(top_bottom_sem_100$media >= mean(tabela_compilada$RESULTADO), "Os cinco melhores", "Os cinco piores")
        
        # Chamar a função grafico_histograma e passar os dados processados
        grafico_histograma(top_bottom_sem_100, tabela_compilada)
      }
      
      # Função para criar um gráfico de histograma com base em dados processados
      grafico_histograma <- function(dados_media, tabela_compilada)
      {
        # Definir cores para os grupos
        cores <- c("Os cinco melhores" = "green", "Os cinco piores" = "red")
        
        # Criar o gráfico de barras
        media_top5_e_bottom5 <- ggplot(dados_media, aes(x = reorder(SETOR, media), y = media, fill = Grupo)) +
          geom_bar(stat = "identity") +
          scale_fill_manual(values = cores) + 
          geom_hline(yintercept = mean(tabela_compilada$RESULTADO), color = "red", linetype = "dashed") + 
          labs(title = "Os cinco setores com as melhores e piores médias em 2023.1",
               x = "Setor",
               y = "Média de Resultado (%)") +
          geom_text(aes(label = sprintf("%.2f", media)), vjust = -0.5, size = 4) + 
          annotate("text", x = 1, y = mean(tabela_compilada$RESULTADO), label = sprintf("%.3f", mean(tabela_compilada$RESULTADO)), vjust = -0.5, color = "red") +
          coord_cartesian(ylim = c(75, 100))
        
        return(media_top5_e_bottom5)
      }
      
      # Função para criar um gráfico de boxplot
      grafico_boxplot <- function(tabela_compilada)
      {
        # Criar o gráfico de boxplot para visualizar a distribuição de resultados por períodos avaliativos
        distribuicao_resultado <- ggplot(tabela_compilada, aes(x = ANO, y = RESULTADO * 100)) + 
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
          geom_hline(yintercept = mean(tabela_compilada$RESULTADO), 
                     linetype = "dashed", color = "red", size = 1) +
          theme(legend.position = "none") +
          coord_cartesian(ylim = c(70, 100))
        
        return(distribuicao_resultado)
      }
      
      # Função para criar um gráfico de linha (série temporal)
      grafico_linha <- function(pessoas_no_grupo)
      {
        nova_cesta <- tabela_compilada %>%
          group_by(ANO, SETOR) %>%
          summarise(media = 100 * mean(RESULTADO)) %>%
          filter(media != 100)
        
        pessoas_no_grupo <- tabela_compilada %>%
          semi_join(nova_cesta, by = c("ANO", "SETOR"))
        
        
        # Calcular média e desvio padrão por ANO
        serie_temporal <- pessoas_no_grupo %>% group_by(ANO) %>% summarise(media = 100 * mean(RESULTADO),
                                                                               desvio_padrao = sd(RESULTADO))
        
        # Criar o gr?fico de linha (série temporal) com média e desvio padrão
        media_e_desvio_padrao_pelo_tempo <- ggplot(serie_temporal, aes(x = ANO, y = media, group = 1)) +
          geom_point(color = "blue", size = 3) +
          geom_line(color = "red") +
          geom_text(aes(label = sprintf("%.3f", desvio_padrao)), vjust = -1, hjust = 0.5, size = 4, color = "black") +
          geom_text(aes(label = sprintf("%.1f", media)), vjust = 2, hjust = 0.4, size = 5, color = "black") +
          labs(x = "Ano", y = "Média de Valores", title = "Série Temporal e Desvio Padrão") +
          theme_classic() + 
          coord_cartesian(ylim = c(94,97))
        
        return(media_e_desvio_padrao_pelo_tempo)
      }
      
      # Esta função calcula as médias e contagens de servidores por SETOR e ANO a partir do quadro de dados 'cesta_de_indicadores'.
      # Retorna um quadro de dados com médias e contagens para cada combinaçõeo de SETOR e ANO.
      calcular_medias_dos_setores_em_cada_avaliacao <- function(cesta_de_indicadores) {
        # Agrupa os dados por SETOR e ANO, calcula a média e a contagem de servidores e cria um quadro de dados resultante.
        media_por_setores <- cesta_de_indicadores %>%
          group_by(SETOR, ANO) %>%
          summarise(media = mean(RESULTADO), servidores = n()) %>%
          pivot_wider(names_from = ANO, values_from = c(media, servidores), names_glue = "{.value}_{ANO}")
        return(media_por_setores)
      }
      
      # Esta função calcula a variabilidade entre as médias de diferentes avaliações (ANO) para cada SETOR.
      # Ela identifica as linhas em que pelo menos duas avaliações têm valores não ausentes (não NA) e calcula a variabilidade entre elas.
      # Retorna um quadro de dados contendo apenas as linhas em que pelo menos duas avaliações têm valores não ausentes (não NA).
      variabilidade_por_setor <- function(media_por_setores) {
        # Filtra as linhas em que pelo menos duas avaliações têm valores não ausentes.
        resultado <- media_por_setores[rowSums(!is.na(media_por_setores[c("media_2022.1", "media_2022.2", "media_2023.1")])) >= 2, ] %>%
          # Calcula a variabilidade entre as avaliações 2022.1 e 2022.2 e entre 2022.2 e 2023.1.
          mutate(variacao_2022.1_2022.2 = ifelse(!is.na(media_2022.1) & !is.na(media_2022.2),
                                                 media_2022.2 - media_2022.1, NA_real_),
                 variacao_2022.2_2023.1 = ifelse(!is.na(media_2022.2) & !is.na(media_2023.1),
                                                 media_2023.1 - media_2022.2, NA_real_)) 
        return(resultado)
      }
      # Função para remover dados com base em critérios
      remover_dados <- function(tabela_compilada) {
        # Filtra o dataframe para o ano de 2022
        tabela_2022 <- tabela_compilada %>% filter(grepl("2022", tabela_compilada$ANO)) %>% distinct()
        # Calcula a ocorrência das matrículas
        ocorrencia_matricula <- table(tabela_2022$MATRICULA)
        matriculas_mais_de_2_ocorrencias <- names(ocorrencia_matricula[ocorrencia_matricula > 2])
        
        # Filtra as matrículas com mais de 2 ocorrências
        esta_na_tabela <- tabela_2022 %>% filter(MATRICULA %in% matriculas_mais_de_2_ocorrencias)
        
        # Remove linhas da tabela original
        tabela_compilada <- anti_join(tabela_compilada, esta_na_tabela[c(1, 3, 8), ])
        
        tabela_2022 <- tabela_compilada %>% filter(grepl("2022", tabela_compilada$ANO)) %>% distinct()

        nomes <- c("CARLOS LEANDRO DOS SANTOS REGINALDO", "LEONARDO DE MELO NOGUEIRA", 
                   "CARLOS ALBERTO ALENCAR FIGUEIREDO DA FROTA", "CRISTINE SIQUEIRA DA SILVA RAPOSO",
                   "PEDRO GIL FERNANDES PINTO")
        
        setores <- c("GP5", "DSI", "GP4", "GP5", "COP")
        
        # retirar as pessoas com esses nomes e setores da tabela_compilada no ano de 2023.1
        tabela_compilada <- tabela_compilada %>%
          filter(!(NOME %in% nomes & SETOR %in% setores & ANO == "2023.1"))

         tabelas <- list(tabela_compilada, tabela_2022)
        
        return(tabelas)
      }
      
      # Função para processar dados
      processar_dados <- function(tabela_compilada, tabela_2022) {
        # Calcula médias
        
        medias_2022 <- tabela_2022 %>%
          group_by(MATRICULA) %>%
          summarise(across(c(`NOTA QUANTITATIVA`, `NOTA QUALITATIVA`, RESULTADO), mean)) %>%
          left_join(filter(select(tabela_2022, c(1,3:6)), ANO == "2022.2"), by = "MATRICULA") %>% distinct()
        
        # Remove colunas indesejadas
        medias_2022 <- select(medias_2022, -5)
        
        # Preenche valores NA em SETOR e ORGÃO 1º NÍVEL
        medias_2022 <- medias_2022 %>%
          mutate(
            SETOR = ifelse(is.na(SETOR), tabela_compilada$SETOR, SETOR),
            `ÓRGÃO 1º NÍVEL` = ifelse(is.na(`ÓRGÃO 1º NÍVEL`), tabela_compilada$`ÓRGÃO 1º NÍVEL`, `ÓRGÃO 1º NÍVEL`),
            NOME = ifelse(is.na(NOME), tabela_compilada$NOME, NOME)
          )
        
        # Atribui Nivel_Desempenho com base nos resultados
        medias_2022 <- medias_2022 %>%
          mutate(
            Nivel_Desempenho = case_when(
              RESULTADO <= 0.2999 ~ "Não atende",
              RESULTADO >= 0.3 & RESULTADO <= 0.6999 ~ "Atendimento parcial",
              RESULTADO >= 0.7 & RESULTADO <= 0.8999 ~ "Atendimento pleno",
              RESULTADO >= 0.9 & RESULTADO <= 0.9999 ~ "Atendimento com excelência",
              RESULTADO == 1 ~ "Destaque",
              TRUE ~ "Outro"
            )
          )
        
        # Seleciona as colunas desejadas
        medias_2022 <- select(medias_2022, 1, 7, 2, 3, 4, 8, 5, 6)
        
        media_qualitativo_2022 <- tabela_2022 %>%
          group_by(MATRICULA) %>%
          summarize(across(12:30, ~mean(., na.rm = TRUE))) 
        
        medias_2022 <- full_join(medias_2022, media_qualitativo_2022, by = "MATRICULA") 
        
        # Salva os resultados em um arquivo Excel
        write_xlsx(medias_2022, "medias_2022.xlsx")
        
        # Retorna o dataframe processado
        return(medias_2022)
      }

      # Função para calcular e escrever a média dos fatores qualitativos por setor em um arquivo Excel
      calcular_media_qualitativo_por_setor <- function(tabela_compilada)
      {
        # Agrupar por ano (ANO) e setor (SETOR), calcular a média para as colunas 11 a 29
        media_qualitativo_por_setor <- tabela_compilada %>% 
          group_by(ANO, SETOR) %>% 
          summarise(across(11:29, ~mean(., na.rm = TRUE)))
        
        # Escrever o resultado em um arquivo Excel
        write_xlsx(media_qualitativo_por_setor, "F:/CDP/Avaliação de Desempenho/Ciclo 2023/media_fatores_qualitativos_por_setor.xlsx")
        
        # Retornar os valores calculados
        return(media_qualitativo_por_setor)
      }
      
      # Função para calcular e escrever a média dos fatores qualitativos por ano em um arquivo Excel
      calcular_media_por_qualitativo <- function(tabela_compilada)
      {
        # Agrupar por ano (ANO), calcular a média para as colunas 12 a 30
        media_por_qualitativo <- tabela_compilada %>% 
          group_by(ANO) %>% 
          summarise(across(12:30, ~mean(., na.rm = TRUE)))
        
        # Calcular a média por linha para cada coluna, excluindo a primeira coluna (ANO)
        for (i in 2:length(media_por_qualitativo)) {
          media_por_qualitativo[4, i] <- sum(media_por_qualitativo[, i], na.rm = TRUE) / sum(!is.na(media_por_qualitativo[, i]))
        }
        
        # Ordenar as colunas com base na média calculada por linha
        last_vals <- unlist(media_por_qualitativo[nrow(media_por_qualitativo), -1])
        media_por_qualitativo <- media_por_qualitativo[, c("ANO", names(sort(last_vals)))]      
        
        # Escrever o resultado em um arquivo Excel
        write_xlsx(media_por_qualitativo, "F:/CDP/Avaliação de Desempenho/Ciclo 2023/media_fatores_qualitativos_por_ano.xlsx")
        
        # Retornar os valores calculados
        return(media_por_qualitativo)
      }
      
      # Função para calcular a média dos fatores qualitativos por perfil
      calcular_media_qualitativo_por_perfil <- function(cesta_de_indicadores)
      {
        # Agrupar por ano (ANO) e perfil (PERFIL), calcular a média para `NOTA QUALITATIVA`
        media_qualitativo_por_perfil <- cesta_de_indicadores %>% 
          group_by(ANO, PERFIL) %>% 
          summarise(media = mean(`NOTA QUALITATIVA`))
        
        # Retornar os valores calculados
        return(media_qualitativo_por_perfil)
      }
      
      # Definir caminhos e filtros para buscar os arquivos
      caminho <- c("F:/CDP/Avaliação de Desempenho/Ciclo 2022/BACKUP","M:/")
      filtro_av1 <- c("Banco|Banco2|Servidor|~\\$|.pdf|_MATRIZ PRODUTOS|MODELO|TabelaCompilada|AV_Bimestral_2|Avaliação 1º ciclo|Pactuação 2º Ciclo|202209-202208|Matriz|AV_2_Bimestral|GC6preenchida|AV_Bimesrtal_2|STE-30-05-2023|M://GAP/2022/|Pactuacao_AV_2|PARA AVALIAÇÃO SUBJETIVA")
      filtro_av2 <- c("~\\$|historico|pdf|PDF|_MATRIZ PRODUTOS|MODELO|TabelaCompilada|GC6preenchida|Tabela_Compilada|Matriz_produtos|Tabela Compilada|.txt|M://SUBPES/AV_Bimestral_2/SUB-SEGURIDADE/")
      
      # Abrir pastas relacionadas Ã¡ AV_Bimestral_1 e AV_Bimestral_2
      pastas_av1 <- abrir_av_bimestral_1(caminho, filtro_av1)
      pastas_av2 <- abrir_av_bimestral_2(caminho, filtro_av2)
      
      # Ler e processar os dados dos arquivos da AV_Bimestral_1 e AV_Bimestral_2
      tabela_av1 <- map(pastas_av1$arquivo, resultado_Av1) %>% bind_rows()
      tabela_av2 <- map(pastas_av2$arquivo, resultado_Av2) %>% bind_rows()
      
     
      # Combinar as tabelas de AV_Bimestral_1 e AV_Bimestral_2 e realizar limpeza dos dados
      tabela_compilada <- full_join(tabela_av1, tabela_av2) %>% limpeza_dos_dados() 
      
      tabela_compilada <- transformacao_dados(tabela_compilada)
      
      # Chamada da função para remover as avaliaÃ§Ãµes erradas
      tabelas <- remover_dados(tabela_compilada)
      
      tabela_compilada <- tabelas[[1]]
      
      tabela_2022 <- tabelas[[2]]
      
      # Criada a tabela com as Médias quantitativa, qualitativa e o resultado de cada servidor em 2022
      media_2022 <- processar_dados(tabela_compilada, tabela_2022)
      
      # Criar um arquivo Excel com os dados processados
      #criar_excel(tabela_compilada)
      
      # Gerar gráfico de histograma com os cinco melhores e cinco piores setores
      histograma_top5_e_bottom5 <- valores_grafico_histograma(tabela_compilada)
      
      # Gerar gráfico de boxplot para visualizar a distribuição de resultados por períodos avaliativos
      metricas_dos_resultados <- grafico_boxplot(tabela_compilada)
      
      # Gerar gráfico de linha (série temporal) com médias e desvio padrão
      serie_temporal_resultado_total <- grafico_linha(pessoas_no_grupo)
      
      # Exibir os gráficos
      plot(histograma_top5_e_bottom5)
      plot(metricas_dos_resultados)
      plot(serie_temporal_resultado_total)
      
      # Chama a função 'calcular_medias_dos_setores_em_cada_avaliacao' para calcular médias e contagens de servidores por SETOR e ANO.
      media_por_setores <- calcular_medias_dos_setores_em_cada_avaliacao(tabela_compilada)
      
      # Chama a função 'variabilidade_por_setor' para calcular a variabilidade entre as médias de diferentes avaliações (ANO) para cada SETOR.
      resultado <- variabilidade_por_setor(media_por_setores)
      
      # Chama essa função calcular a média de cada nota qualitativa nos setores em todos os períodos avaliativos
      media_qualitativo_por_setor <- calcular_media_qualitativo_por_setor(tabela_compilada)

      # Chama essa função calcular a media da nota qualitativa para cada perfil de colaborador em todos os períodos avaliativos
      media_qualitativo_por_perfil <- calcular_media_qualitativo_por_perfil(tabela_compilada)
      
      # Chama essa função para calcular a média de cada nota qualitativa nos períodos avaliativos 
      media_qualitativo <- calcular_media_por_qualitativo(tabela_compilada)

    
      
