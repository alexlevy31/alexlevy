# Handoff: búsqueda de diamante de compromiso (óvalo natural)

> Para Claude Code: este archivo resume todo lo investigado en una sesión anterior.
> Léelo completo antes de seguir. Responde en español, en tono casual y claro.
> El usuario no es experto en diamantes: explica términos sin tecnicismos.

## Contexto
- El usuario está en **Panamá** y busca un **diamante ovalado natural** para un anillo de compromiso.
- Presupuesto inicial: **US$8,000–10,000**. Puede estirarse si vale la pena.
- Hay dos vendedores locales (en el Excel se llaman "Vendedor 1" y "Vendedor 2"; **no usar nombres de personas**).
- Todo lo de los vendedores es **por pedido**: sacan las piedras de las mismas bases de datos mayoristas que se publican en línea.
- Los precios del Vendedor 2 ya incluyen el descuento de $1,300 (precio de lista − $1,300). **Son precios finales.**

## Archivos del repo
- `Comparacion_Diamantes.xlsx`: el Excel principal, con 6 pestañas:
  - **Comparación**: las 10 piedras ofrecidas en Panamá, con ranking, medidas, área visible, $/mm² y rango de mercado.
  - **Resumen research**: ranking y hallazgos clave.
  - **CaratHunter**: alternativas en línea con certificado GIA, medidas (celdas amarillas a llenar), área estimada y real, y links.
  - **Qué revisar**: guía de qué mirar en un diamante.
  - **Precios de mercado**: referencias con links.
  - **Pendientes**: checklist.
- `scripts/build_excel.py`: genera el Excel desde cero (`python scripts/build_excel.py`). **Para cambiar el Excel, edita este script y vuelve a correrlo**; no edites el .xlsx a mano. Requiere `openpyxl`.

## Criterios usados (óvalo)
- **Área visible** = π/4 × largo × ancho. Es lo que define qué tan grande se ve. Comparar por mm, **no por quilates**.
- Profundidad ideal **58–64%** (alto/ancho). Más de ~65% = peso escondido, se ve más chico. Menos de 60% = más riesgo de bow-tie o "window".
- Largo/ancho **1.35–1.50** (clásico). Mesa 53–63%. Girdle de thin a slightly thick o thick (nada "extremely thick").
- Color G–I alcanza. Claridad SI1 o SI2 que sea eye-clean. Evitar I1.
- Fluorescencia None, Faint o Medium. Strong solo con descuento y viéndola en persona.
- **GIA** es el laboratorio preferido. **HRD** suele calificar 1–2 grados más alto y vale bastante menos.
- Referencia "100%" para comparar tamaños: el #2 de Panamá, **10.36 × 7.05 = 57.4 mm²**.

## Piedras de Panamá (precios finales)
| Ranking | Piedra | Lab | Medidas | Precio | Veredicto |
|---|---|---|---|---|---|
| 1 | 1.70 I SI1, VG/Ex, Faint | GIA | 8.96×6.86×4.35 | $7,000 | Precio justo, la de menos riesgo |
| 2 | 2.01 H SI2, Ex/Ex | GIA | 10.36×7.05×4.32 | $12,600 | La mejor de 2 ct en Panamá. Falta el PDF GIA y la fluorescencia |
| 3 | 2.01 G SI2, Ex/VG | GIA | 9.95×7.18×4.46 | $13,700 | Reemplazo del #2 |
| 4 | 2.01 E I1, VG/Ex, None | GIA 7558507203 | 9.68×6.84×4.47 | $11,700 | I1, girdle extremely thick. En línea desde $9,150 |
| 5 | 1.85 D SI2, Ex/Ex, None | **HRD** 250000260720 | 9.93×6.79×3.96 | $10,500 | Caro por ser HRD. En línea desde $9,001. Inscripción del girdle desconocida |
| 6 | 2.00 I SI1, Ex/Good, Strong Blue | GIA | 9.29×7.32×4.58 | $10,500 | Strong Blue, muy redondo |
| 7 | 2.01 H SI2 "alargado", Ex/VG, Faint | GIA | 10.69×6.40×4.25 | $12,600 | Profundo (66%), se ve más chico que el #2 |
| 8 | 1.71 H SI2, Ex/Ex, None | GIA (…8893) | 9.38×6.51×4.20 | $8,600 | Inclusiones en el centro, girdle grueso |
| 9 | 2.01 H SI2, Ex/Ex, None | GIA (…2014) | 9.20×6.78×4.72 | $11,400 | 69.7% de profundidad: se ve como un 1.70 |
| — | 2.51 I SI2 | GIA | 10.59×7.58×4.81 | — | No disponible |

## Hallazgo principal: en línea es bastante más barato
- **CaratHunter** (carathunter.com) es un comparador: la misma piedra (mismo GIA) aparece en muchas tiendas con precios muy distintos. "Desde" es la tienda más barata.
- El usuario **compró el pase de CaratHunter** (~US$27 / 30 días), así que **él sí puede ver los nombres de las tiendas**. Desde la nube, las páginas de detalle tienen verificación anti-bots (Cloudflare). **No intentar saltarla.** Pedirle capturas al usuario.
- La búsqueda de CaratHunter sí se puede leer sin JS (`/us/advanced-search?origin=natural&shape=OVAL&carat_min=..&carat_max=..&colors=G,H,I&clarities=SI1,SI2&labs=GIA&sort=price_asc`). Trae JSON embebido con carat, color, claridad, polish, simetría, fluorescencia, ratio, `trusted_min_price`, `certificate_number` y `id`, pero **no trae medidas ni tienda**.

### Favorito actual: 2.01 G SI1, GIA 2567819843 ✅
- 9.98 × 7.15 × 4.34 mm, **profundidad 60.7%**, mesa 63%, L/W 1.40, girdle slightly thick–thick, culet none, Ex/VG/None. Reporte del 2-nov-2026.
- Comentarios del GIA: "Clouds are not shown. Pinpoints are not shown."
- Mapa de inclusiones: **grupo de cristales en la mesa, un poco abajo del centro** (hay que confirmar en video que no se vea), una aguja arriba, y feathers chicas en las orillas.
- Tiendas: Ting Diamond (Hong Kong) $10,330 · Sarvada $10,951 · **DIANA Jewellery (Dubái) $11,280** · **Brilliant Earth $12,260** · … hasta $33,097 (25 tiendas).
- Ting: **no envía a Panamá** (solo a UK, AU, CA y US), pide 60% de adelanto y está pensada para comprar en su tienda de Hong Kong.
- DIANA: envía internacional asegurado, pero **las piedras certificadas no tienen devolución** y hay que pagar todo antes del envío.
- **Recomendación dada:** pedirle a un joyero de Panamá que consiga este GIA exacto (objetivo ~$11,000–11,500). Si no, Brilliant Earth ($12,260, con 30 días de devolución y pago con tarjeta).

### Otras alternativas en línea (GIA, naturales; los precios son "desde")
Están en la pestaña CaratHunter, con links. Las principales:
- 2.01 H SI2 Ex/Ex None, GIA 6521053346, desde $7,685
- 2.01 G SI2 Ex/Ex None, GIA 1538651314, desde $7,457
- 2.01 H SI2 Ex/VG None, GIA 7548477701, desde $6,805
- 1.80 I SI1 Ex/Ex None, GIA 7558423822, desde $6,936
- 1.71 I SI1 Ex/Ex None, GIA 2524515119, desde $5,259
- ❌ 1.70 I SI1, GIA 1569498556: está en Brian Gavin a $5,418, pero tiene **68.8% de profundidad** (8.88×6.33×4.35). Descartada.

### Brilliant Earth (listas indexadas, de ~50 días atrás; pueden estar desactualizadas)
- **2.01 G SI2 "Super Ideal" a $6,820**: muy prometedor y sin verificar. El usuario quedó en buscarlo en el sitio (filtro natural, 2.00–2.10, G, SI2, Super Ideal, ordenado por precio).
- 1.93 I SI2 Ideal $5,150 · 1.90 G SI1 Super Ideal $12,980 · 2.01 G SI2 Super Ideal $11,610.
- Brilliant Earth bloquea el scraping desde la nube. En el PC del usuario, pedirle el link o una captura.

## Comprar fuera y traerlo a Panamá
- Muchas tiendas de EE.UU. **no envían a Panamá**. Opciones: mandarlo a una dirección en EE.UU. y traerlo en un viaje, un casillero o courier que acepte joyas de ese valor con seguro, o comprarlo en persona.
- Aduana de Panamá: arancel general de hasta ~15% sobre el valor CIF, más 7% de ITBMS, y posiblemente un impuesto selectivo a joyas. **La tasa exacta para diamantes sueltos no está confirmada**: verificar en aranceles.ana.gob.pa o con el courier.
- Siempre: pagar con tarjeta de crédito, devolución por escrito, envío asegurado, y verificar que la inscripción láser coincida con el número GIA.

## Pendientes / siguientes pasos
1. Buscar en Brilliant Earth el **2.01 G SI2 Super Ideal ~$6,820** y revisar su GIA (medidas, profundidad, mapa).
2. Pedir **video** del 2.01 G SI1 (GIA 2567819843) para ver el grupo de cristales.
3. Con el pase de CaratHunter, abrir las alternativas de 2.01 H/G SI2 y 1.80 I SI1, y anotar en el Excel las medidas reales (celdas amarillas) y las tiendas.
4. Pedir a los joyeros de Panamá cotización de los certificados encontrados en línea.
5. Del #2 de Panamá (2.01 H SI2, 10.36×7.05): falta el PDF del GIA y la fluorescencia.

## Notas técnicas
- `Comparacion_Diamantes.xlsx` se regenera con `python scripts/build_excel.py`. Las fórmulas se calculan al abrir en Excel.
- Para revisar un mapa de inclusiones de un PDF de GIA: renderizar a 300 dpi (`pdftoppm -r 300`) y recortar la zona de "Clarity Characteristics".
- GIA report check (gia.edu) tiene anti-bots: que el usuario descargue el PDF y lo comparta.
