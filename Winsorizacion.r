# -----------------------------------------------------------------------------
#  Winsorización usando lista de outliers y Boxplots
# -----------------------------------------------------------------------------
# Descripción:
# Este script carga dos archivos Excel:
#  - Base completa de datos.
#  - Lista de outliers (con columnas: 'columna' y 'fila').
# Luego aplica winsorización únicamente en las celdas indicadas,
# y genera boxplots antes y después para todas las variables numéricas.
# Finalmente guarda la base winsorizada en un nuevo Excel.

# --- 1. Instalación y carga de paquetes ---
if (!requireNamespace("readxl", quietly = TRUE)) install.packages("readxl")
if (!requireNamespace("writexl", quietly = TRUE)) install.packages("writexl")
if (!requireNamespace("dplyr", quietly = TRUE)) install.packages("dplyr")
if (!requireNamespace("DescTools", quietly = TRUE)) install.packages("DescTools")
if (!requireNamespace("tidyr", quietly = TRUE)) install.packages("tidyr")
if (!requireNamespace("ggplot2", quietly = TRUE)) install.packages("ggplot2")

library(readxl)
library(writexl)
library(dplyr)
library(DescTools)
library(tidyr)
library(ggplot2)

# --- 2. Parámetros de entrada/salida y winsorización ---
ruta_base     <- "C:/Users/nicot/OneDrive/Escritorio/Trabajo_Inferencia/Limpieza_Base/archivo_limpio.xlsx"# Base de datos completa
ruta_outliers <- "C:/Users/nicot/OneDrive/Escritorio/Trabajo_Inferencia/Limpieza_Base/outliers_detalle.xlsx"# Excel con columnas 'columna', 'fila'
ruta_salida   <- "C:/Users/nicot/OneDrive/Escritorio/Trabajo_Inferencia/Limpieza_Base/base_winsorizada.xlsx"# Archivo Excel de salida
lower_prob    <- 0.05# Percentil inferior
upper_prob    <- 0.95# Percentil superior

# --- 3. Cargar datos ---
df_base     <- read_excel(path = ruta_base)
df_outliers <- read_excel(path = ruta_outliers)

# --- 4. Gráfico antes de winsorización ---
df_long <- df_base %>% select(where(is.numeric)) %>%
  pivot_longer(everything(), names_to = "variable", values_to = "value")

p1 <- ggplot(df_long, aes(x = variable, y = value)) +
  geom_boxplot() +
  ggtitle("Boxplot antes de winsorización") +
  theme(axis.text.x = element_text(angle = 45, hjust = 1))
print(p1)

# --- 5. Calcular límites por variable ---
limits <- df_base %>%
  summarise(across(
    where(is.numeric),
    list(
      lower = ~ quantile(.x, lower_prob, na.rm = TRUE),
      upper = ~ quantile(.x, upper_prob, na.rm = TRUE)
    ), .names = "{col}_{fn}"))

# Convertir limits a formato largo para join
limits_long <- limits %>%
  pivot_longer(everything(), names_to = c("variable","bound"),
               names_sep = "_", values_to = "threshold")

# --- 6. Aplicar winsorización solo en celdas indicadas ---
df_wins <- df_base
for(i in seq_len(nrow(df_outliers))) {
  col <- df_outliers$columna[i]
  row <- df_outliers$fila[i]
  lims <- limits_long %>% filter(variable == col)
  lower <- lims$threshold[lims$bound == "lower"]
  upper <- lims$threshold[lims$bound == "upper"]
  val <- df_wins[[col]][row]
  # Reemplazar si está fuera de límites
  if(!is.na(val)) {
    df_wins[[col]][row] <- min(max(val, lower), upper)
  }
}

# --- 7. Gráfico después de winsorización ---
df_wins_long <- df_wins %>% select(where(is.numeric)) %>%
  pivot_longer(everything(), names_to = "variable", values_to = "value")

p2 <- ggplot(df_wins_long, aes(x = variable, y = value)) +
  geom_boxplot() +
  ggtitle("Boxplot después de winsorización") +
  theme(axis.text.x = element_text(angle = 45, hjust = 1))
print(p2)

# --- 8. Guardar datos winsorizados ---
write_xlsx(df_wins, path = ruta_salida)
cat("\nBase winsorizada guardada en: ", ruta_salida, "\n")
# -----------------------------------------------------------------------------
# Fin del script
# -----------------------------------------------------------------------------
