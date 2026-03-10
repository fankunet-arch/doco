#!/usr/bin/env python3
"""Translate 墙贴提示_V1_22.docx to Spanish, preserving formatting."""

from docx import Document
from docx.oxml.ns import qn
import copy
import re

doc = Document('墙贴提示_V1_22.docx')
body = doc.element.body

# Cell-level translation map: (box_index_mod, table, row, col) -> translation
# Since boxes come in pairs (0&1, 2&3, etc.), we use box_index % 2 == 0 pattern
# and apply same translations to both copies.

# We'll translate by matching the Chinese cell text to Spanish.
# This mapping covers all unique cell texts across all text boxes.

cell_translations = {
    # === Box 0/1: 咸奶酪奶盖 (Salty Cheese Foam) ===
    '咸奶酪奶盖': 'Espuma de Queso Salada',
    '奶油奶酪': 'Queso crema',
    '纯牛奶': 'Leche entera',
    '淡奶油': 'Nata',
    '白糖': 'Azúcar',
    '小份': 'Peq.',
    '大份': 'Grande',

    # Instructions box 0/1
    '打蛋器铁锅中加奶油奶酪、冰纯牛奶，最低速打约25秒搅匀后，依次加入淡奶油、炼乳搅拌均匀放入打蛋器上快速打发。打发至状态层层叠起，缓缓回落3-5秒':
        'En olla de batidora, añadir queso crema y leche fría. Batir a vel. mín. ~25s hasta mezclar. Añadir nata y leche condensada, mezclar bien. Batir a vel. alta. Batir hasta que forme capas y caiga lentamente en 3-5s.',

    # === Box 2/3: 西米 (Sagú) ===
    '西米': 'Sagú',
    '热水': 'Agua cal.',
    '直饮水': 'Agua pot.',
    '糖浆': 'Jarabe',
    '配比一': 'Prop. 1',
    '配比二': 'Prop. 2',
    '配比三': 'Prop. 3',
    '配比四': 'Prop. 4',
    '50克': '50g',
    '100克': '100g',
    '200克': '200g',

    '珍珠锅加热水按②键流程烧水加西米煮制。用直饮水和糖浆配制糖水。煮制完成后取出沥水后立即用可饮用的冷水冲洗降温，倒入糖水中备用。':
        'Olla de perlas: añadir agua cal., pulsar ② para hervir, añadir sagú y cocinar. Preparar agua azucarada con agua pot. y jarabe. Al terminar, escurrir e inmediatamente enjuagar con agua fría potable. Verter en agua azucarada y reservar.',

    # === Box 4/5: 血糯米 (Arroz glutinoso negro) ===
    '血糯米': 'Arroz glut. negro',
    '白砂糖': 'Azúcar blanca',
    '煮制时间': 'T. cocción',
    '1500 ml': '1500 ml',
    '400 g': '400 g',
    '100 g': '100 g',
    '30分钟': '30 min',
    '3000 ml': '3000 ml',
    '800 g': '800 g',
    '200 g': '200 g',

    '取血糯米清洗搓洗2次，清洗干净。高压锅加直饮水+血糯米，电磁炉按键开始煮豆，代表机器已经开始运作煮制时间，无需其他操作。煮制时间到后，将秤好的白砂糖倒入煮好血糯米锅中搅拌白砂糖完全融化至粘稠度后按保温键保温备用。':
        'Lavar arroz glut. negro frotando 2 veces hasta limpiar. Olla a presión: añadir agua pot. + arroz, pulsar inicio cocción legumbres (máquina opera automát., sin intervención). Al terminar, añadir azúcar blanca pesada, revolver hasta disolver completamente y lograr consistencia espesa. Pulsar mantener caliente y reservar.',

    # === Box 6/7: 抹茶液 (Líquido matcha) ===
    '抹茶液': 'Líq. matcha',
    '抹茶粉': 'Polvo matcha',
    # '淡奶油' already mapped
    # '直饮水' already mapped

    '操作流程：\xa0取冰沙机杯先加直饮水、淡奶油、抹茶粉、盖盖放冰沙机打约40秒；打好后倒入保险壶放冰箱冷藏':
        'Proceso: En vaso de licuadora, añadir agua pot., nata y polvo matcha. Tapar y licuar ~40s. Verter en termo y refrigerar.',

    # === Box 8/9: 2号滇红 (Té Dianhong #2) ===
    '2号滇红': 'Té Dianhong #2',
    '热水(95℃+)': 'Agua cal. (95°C+)',
    '茶叶': 'Hojas de té',
    '泡制计时': 'T. infusión',
    '冰块': 'Hielo',
    '小量': 'Peq.',
    '大量': 'Grande',
    '50克': '50g',
    '100克': '100g',
    '10分钟': '10 min',
    '1000克': '1000g',
    '2000克': '2000g',

    '流程：(取5L量桶)→量取热水→(单独量茶叶)倒茶叶→吧勺顺时针搅10圈→加盖(实盖) →泡制计时→秤取冰块倒入保温茶桶备用→时间到，过滤茶汤至茶桶(尾段不要)→搅匀':
        'Proceso: (Tomar jarra 5L) → medir agua cal. → (medir hojas aparte) añadir hojas → revolver 10 vueltas en sentido horario → tapar (tapa sólida) → cronometrar infusión → pesar hielo y verter en termo → al terminar, filtrar té al termo (descartar final) → mezclar.',

    # === Box 10/11: 茉莉绿茶 (Té verde jazmín) ===
    '茉莉绿茶': 'Té Verde Jazmín',
    '热水(75℃+)': 'Agua cal. (75°C+)',
    '6分钟': '6 min',
    '145克': '145g',
    '290克': '290g',
    '750克': '750g',
    '1500克': '1500g',

    '(取1L量杯)称取冰块备用→(取5L量桶)量取热水→热水中加入降温冰→倒茶叶→吧勺顺时针搅3-4圈→不加盖→计时→计时到前，秤取冰块倒入保温茶桶备用→时间到，过滤茶汤至茶桶(尾段不要)→吧勺搅匀':
        '(Tomar jarra 1L) pesar hielo y reservar → (tomar jarra 5L) medir agua cal. → añadir hielo de enfriamiento → añadir hojas → revolver 3-4 vueltas en sentido horario → NO tapar → cronometrar → antes de terminar, pesar hielo y verter en termo → al terminar, filtrar té al termo (descartar final) → mezclar.',

    # === Box 12/13: 冰水 (Agua helada) ===
    '冰水': 'Agua Helada',
    '饮用水': 'Agua pot.',
    '早上进店': 'Apertura',
    '中途补水': 'Reposición',
    '至4L': 'hasta 4L',

    # === Box 14/15: 黑糖珍珠 (Perlas de azúcar moreno) ===
    '黑糖珍珠': 'Perlas Azúcar Moreno',
    '珍珠': 'Perlas',
    '黑糖粉': 'Azúcar mor. polvo',
    '煮珍珠': 'Cocinar',
    '焖珍珠': 'Reposar',
    '炒珍珠': 'Saltear',
    '5分钟': '5 min',

    '将装有热水的内锅放入珍珠锅里，按一下【①】号键亮灯再点击【开始】按键后，【开始】按键会亮红灯，代表机器已经开始运作。'
    '水烧开沸腾鸣笛，看到【加珍珠】按键红灯发亮时，先把需要加的珍珠量称取后倒入过筛网中，在水槽轻轻抖动滤掉粉末。'
    '打开锅盖加入对应的珍珠量，用黑色珍珠勺搅拌30下，必须确认珍珠已全部搅开，避免黏成糊状。'
    '加完珍珠后，→【煮珍珠】隔5分钟搅拌一次，直到【焖珍珠】无需其他操作，按键灯熄灭后，发出鸣笛声，打开锅盖，搅拌均匀把锅里的水滤掉加入对应量的黑糖粉，搅拌黑糖粉全部搅融化后，手动长按开始键3秒闪亮，按→【进度调整键】至【煮水键】→按【开始键】边煮边搅拌5分钟黑糖粉完全融化至粘稠度。'
    '再次长按【开始键】3秒闪亮按【进度调整键】至【焖珍珠】设置时间6小时保温，隔半小时搅拌一次。':
        'Colocar olla interior con agua cal. en olla de perlas. Pulsar【①】(luz encendida), luego【Inicio】(luz roja = máquina en marcha). '
        'Al hervir (pitido), cuando luz roja de【Añadir perlas】se encienda, pesar perlas, verter en colador y agitar suavemente para quitar polvo. '
        'Abrir tapa, añadir perlas, revolver 30 veces con cuchara negra. Confirmar que todas las perlas estén separadas (evitar grumos). '
        'Después de añadir: →【Cocinar】revolver cada 5 min →【Reposar】sin intervención. Al apagarse luz y sonar pitido, abrir tapa, mezclar bien, escurrir agua, añadir azúcar mor. polvo correspondiente. Revolver hasta disolver. '
        'Mantener pulsado【Inicio】3s (parpadeo) →【Ajuste progreso】hasta【Hervir agua】→【Inicio】cocinar y revolver 5 min hasta consistencia espesa. '
        'Mantener pulsado【Inicio】3s →【Ajuste progreso】hasta【Reposar perlas】. Configurar 6h mantener caliente. Revolver cada 30 min.',

    # === Box 16/17: 布蕾蛋糕酱 (Salsa Crème Brûlée) ===
    '布蕾蛋糕酱': 'Salsa Crème Brûlée',
    '布蕾粉': 'Polvo brûlée',
    # '淡奶油' and '纯牛奶' already mapped

    '\xa0打蛋器铁锅加入布蕾粉、淡奶油、纯牛奶，搅拌均匀放入打蛋器上快速打至绵绸状态即可':
        'En olla de batidora, añadir polvo brûlée, nata y leche. Mezclar bien y batir a vel. alta hasta obtener textura suave y cremosa.',

    # === Box 18/19: 开心果芝士奶盖 (Espuma de queso pistacho) ===
    '开心果芝士奶盖': 'Espuma Queso Pistacho',
    # '奶油奶酪', '纯牛奶' already mapped
    '开心果酱': 'Pasta pistacho',
    'NATA': 'NATA',
    # '白糖' already mapped

    '① 将切好的芝士和开心果酱放入打蛋机锅，用1-2档底速搅拌1分钟后，用刮刀刮一下周边和底部，用3-4档中速搅拌1分钟后转高速搅拌1-2分钟，打至芝士蓬松。'
    '② 将纯牛奶和糖混合搅拌均匀，打蛋机调至2档，缓慢加入纯牛奶，至搅拌均匀后，倒出至量杯中，备用。'
    '③ 将淡奶油倒入打蛋机，用1-2档底速搅拌1分钟后，用3-4档中速搅拌1分钟后转高速搅拌30秒-1分钟。'
    '④ 打蛋机调至2档，将芝士液缓慢倒入奶油中，搅拌均匀即可。'
    '⑤ 将打好的芝士奶盖，倒入量杯中，轻震动几下，密封冷藏保存。':
        '① Queso cortado + pasta pistacho en batidora. Vel. 1-2 por 1 min, raspar bordes y fondo, vel. 3-4 por 1 min, vel. alta 1-2 min hasta queso esponjoso. '
        '② Mezclar leche + azúcar. Batidora vel. 2, añadir leche lentamente. Mezclar bien, verter en jarra y reservar. '
        '③ Nata en batidora. Vel. 1-2 por 1 min, vel. 3-4 por 1 min, vel. alta 30s-1 min. '
        '④ Batidora vel. 2, verter líq. de queso lentamente en nata. Mezclar bien. '
        '⑤ Verter espuma en jarra, agitar suavemente, sellar y refrigerar.',
}


def set_cell_text(cell, new_text):
    """Replace all text in a cell with new_text, preserving first run's formatting."""
    # Get all w:t elements in the cell
    all_t = cell.findall('.//' + qn('w:t'))
    if not all_t:
        return

    # Set the full text on the first w:t element
    all_t[0].text = new_text
    # Preserve spaces
    all_t[0].set(qn('xml:space'), 'preserve')

    # Clear all other w:t elements
    for t in all_t[1:]:
        t.text = ''


txboxes = body.findall('.//' + qn('w:txbxContent'))

for txbox in txboxes:
    tables = txbox.findall(qn('w:tbl'))
    for tbl in tables:
        rows = tbl.findall(qn('w:tr'))
        for row in rows:
            cells = row.findall(qn('w:tc'))
            for cell in cells:
                # Get full cell text
                all_t = cell.findall('.//' + qn('w:t'))
                cell_text = ''.join(t.text or '' for t in all_t)

                if cell_text in cell_translations:
                    set_cell_text(cell, cell_translations[cell_text])

output_file = '墙贴提示_V1_22_ES.docx'
doc.save(output_file)
print(f'Saved translated file as: {output_file}')
