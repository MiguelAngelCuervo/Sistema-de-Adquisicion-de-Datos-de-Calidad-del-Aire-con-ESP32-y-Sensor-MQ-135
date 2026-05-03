#include <stdio.h>
#include "freertos/FreeRTOS.h"
#include "freertos/task.h"
#include "esp_adc/adc_oneshot.h"

// Definimos el canal para el GPIO 34
#define MQ135_CHANNEL ADC_CHANNEL_6 

void app_main(void)
{
    // 1. Configurar la unidad ADC1
    adc_oneshot_unit_handle_t adc1_handle;
    adc_oneshot_unit_init_cfg_t init_config1 = {
        .unit_id = ADC_UNIT_1,
    };
    adc_oneshot_new_unit(&init_config1, &adc1_handle);

    // 2. Configurar el canal del sensor MQ-135 (GPIO 34)
    adc_oneshot_chan_cfg_t config = {
        .bitwidth = ADC_BITWIDTH_DEFAULT, // Resolución de 12 bits (0-4095)
        .atten = ADC_ATTEN_DB_12,         // Atenuación para leer hasta ~3.1V
    };
    adc_oneshot_config_channel(adc1_handle, MQ135_CHANNEL, &config);

    printf("Iniciando lectura del sensor MQ-135 en GPIO 34...\n");

    while (1) {
        int adc_raw = 0;
        
        // Leer el valor crudo del ADC
        adc_oneshot_read(adc1_handle, MQ135_CHANNEL, &adc_raw);
        
        // --- CAMBIO AQUÍ ---
        // Eliminamos el texto "Lectura Raw..." para que Python reciba solo el dígito
        printf("%d\n", adc_raw); 
        
        // Esperar 100ms o 200ms para que la aguja se vea fluida en Python
        vTaskDelay(pdMS_TO_TICKS(100)); 
    }
}