# Este archivo busca todos los archivos .c en tu carpeta src
FILE(GLOB_RECURSE app_sources ${CMAKE_SOURCE_DIR}/src/*.*)

# Añadimos PRIV_REQUIRES esp_adc para que el compilador encuentre los drivers
idf_component_register(SRCS ${app_sources} PRIV_REQUIRES esp_adc)