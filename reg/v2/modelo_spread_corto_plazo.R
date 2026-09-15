# =====================================================================
# Modelo de spread - Instrumentos de corto plazo en soles, <= 1 año
# Incluye: Certificados de Depósito, Instrumentos de Corto Plazo
#          y Papeles Comerciales
#
# spread = b0 + b1*BTC + b2*log(monto) + b3*no_financiero + b4*rating
#        + b5*tasa_base + dummies de instrumento (vs CD)
#        + dummies de año (vs 2022)
# =====================================================================
# install.packages(c("readxl", "dplyr", "fixest"))
library(readxl)
library(dplyr)
library(fixest)

# ---- 1. Leer la base ------------------------------------------------
raw <- read_excel("data1.xlsm", sheet = "DB-Histórica CP (2)")
names(raw) <- trimws(names(raw))     # quita espacios en los nombres

# ---- 2. Corregir etiquetas ------------------------------------------
# Luz del Sur y CxC Pluzen figuran como "Financiero" en algunas filas.
# "Falabella Perú" es el mismo programa de CDs que Banco Falabella.
base <- raw %>%
  filter(!is.na(Emisión)) %>%
  mutate(
    anio   = as.integer(format(as.Date(Emisión), "%Y")),
    Emisor = trimws(Emisor),
    Sector = ifelse(Emisor %in% c("Luz del Sur", "CxC Pluzen"), "Utilities", Sector),
    Emisor = ifelse(Emisor == "Falabella Perú", "Banco Falabella", Emisor)
  )

# ---- 3. Muestra y variables -----------------------------------------
df <- base %>%
  filter(anio >= 2022,                                # tasa base existe desde 2022
         Instrumento %in% c("Certificados de Depósito",
                            "Instrumento de Corto Plazo",
                            "Papeles Comerciales"),
         Divisa == "PEN",
         Duración <= 1,
         !is.na(`Tasa base`)) %>%
  mutate(
    spread        = (Tasa - `Tasa base`) * 10000,     # en pb
    tasa_base     = `Tasa base` * 100,                # en %
    btc           = BTC,
    monto_imput   = as.integer(is.na(`Ofertado MM`)), # 1 si no hay monto ofertado
    monto         = coalesce(`Ofertado MM`, Asignado),
    log_monto     = log(monto),
    no_financiero = as.integer(Sector != "Financiero"),
    rating        = recode(trimws(RATING), "CP-2" = "CP-1-"),
    rating        = relevel(factor(rating), ref = "CP-1+"),
    instrumento   = recode(Instrumento,
                           "Certificados de Depósito"   = "CD",
                           "Instrumento de Corto Plazo" = "ICP",
                           "Papeles Comerciales"        = "PC"),
    instrumento   = relevel(factor(instrumento), ref = "CD"),
    emisor        = tolower(Emisor)
  ) %>%
  filter(!is.na(spread), !is.na(btc), !is.na(log_monto))

cat("Emisiones:", nrow(df), "| Emisores:", n_distinct(df$emisor), "\n")
print(table(Instrumento = df$instrumento,
            Sector = ifelse(df$no_financiero == 1, "No financiero", "Financiero")))

# ---- 4. Modelos -----------------------------------------------------
# Errores estándar agrupados por emisor
m1 <- feols(spread ~ btc + log_monto + no_financiero + rating,
            data = df, cluster = ~emisor)

m2 <- feols(spread ~ btc + log_monto + no_financiero + rating +
              tasa_base + monto_imput + instrumento + i(anio, ref = 2022),
            data = df, cluster = ~emisor)

etiquetas <- c(
  btc            = "Bid-to-cover",
  log_monto      = "Monto (log)",
  no_financiero  = "Sector no financiero (vs financiero)",
  "ratingCP-1"   = "CP-1 (vs CP-1+)",
  "ratingCP-1-"  = "CP-1- (vs CP-1+)",
  tasa_base      = "Tasa base (%)",
  monto_imput    = "Monto ofertado no disponible",
  instrumentoICP = "Instrumento de Corto Plazo (vs CD)",
  instrumentoPC  = "Papel Comercial (vs CD)",
  "anio::2023"   = "Año 2023 (vs 2022)",
  "anio::2024"   = "Año 2024 (vs 2022)",
  "anio::2025"   = "Año 2025 (vs 2022)",
  "anio::2026"   = "Año 2026 (vs 2022)"
)

etable(m1, m2, headers = c("Sin controles", "Con controles"),
       dict = etiquetas, se.below = TRUE, fitstat = ~ n + r2 + ar2)

# ---- 5. Mismo modelo por periodo ------------------------------------
m_2223 <- feols(spread ~ btc + log_monto + no_financiero + rating + tasa_base +
                  instrumento + i(anio, ref = 2022),
                data = filter(df, anio <= 2023), cluster = ~emisor)

m_2426 <- feols(spread ~ btc + log_monto + no_financiero + rating + tasa_base +
                  monto_imput + instrumento + i(anio, ref = 2024),
                data = filter(df, anio >= 2024), cluster = ~emisor)

etable(m_2223, m_2426, headers = c("2022-2023", "2024-2026"),
       dict = etiquetas, se.below = TRUE, fitstat = ~ n + r2)

# ---- 6. Exportar ----------------------------------------------------
write.csv(coeftable(m2), "coeficientes_modelo_corto_plazo.csv")
