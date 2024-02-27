# Автокоррекция записи посещения (нет скобок у группы, перенос на новую строку)
def autocorrection_visit_record(text):
    res = text.replace(" CS - 202", " (CS-202)").replace(" cs-202", " (CS-202)") \
        .replace(" Cs-202", " (CS-202)") \
        .replace(chr(0xA) + "CS-204", " (CS-204)") \
        .replace(chr(0xA) + "(CS-204)", " (CS-204)") \
        .replace(chr(0xA) + " CS-204", " (CS-204)") \
        .replace(chr(0xA) + "(ИС-106(c))", " (ИС-106(с))") \
        .replace(chr(0xA) + " (CS - 304)", " (CS-304)") 
    return res


# Автокоррекция записи группы для русских и английских букв, пробелов и тире
def autocorrection_group_record(text):
    res = text.replace("(c)", "(с)").replace("(С)", "(с)").replace("(C)", "(с)") \
        .replace(" (", "(").replace("( ", "(").replace(" )", ")").replace("))", ")") \
        .replace("cs", "СS").replace("Cs", "СS").replace("cS", "СS").replace("CS", "СS") \
        .replace("сs", "СS").replace("Сs", "СS").replace("сS", "СS").replace("—", "-") \
        .replace(" -", "-").replace("- ", "-").replace(" - ", "-").replace("СS204", "СS-204")
    return res