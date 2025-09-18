# Assets del certificado

Este directorio debe contener los recursos gráficos utilizados por la plantilla PDF de los certificados. Los ficheros deben subirse manualmente para que la generación del PDF utilice la identidad visual completa.

| Nombre del archivo | Descripción | Recomendaciones |
| ------------------ | ----------- | --------------- |
| `fondo-certificado.png` | Figura roja que actúa como fondo en el margen derecho. | Imagen en formato PNG con transparencia. Ancho recomendado ~350-400 px para A4 apaisado. |
| `lateral-izquierdo.png` | Elemento vertical con la identidad corporativa situado en el margen izquierdo. | Altura suficiente para cubrir todo el alto del documento (A4 apaisado ≈ 595 pt). |
| `pie-firma.png` | Bloque con firma y logotipos que se coloca en la esquina inferior izquierda. | Utiliza transparencia para integrarse correctamente con el fondo. Ancho recomendado ~220-260 px. |

> Los nombres de archivo deben coincidir exactamente con los listados arriba. El generador aplica una imagen transparente de reserva si no encuentra alguno de los recursos para evitar errores en tiempo de ejecución.
