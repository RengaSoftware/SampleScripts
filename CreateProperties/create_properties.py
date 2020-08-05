import json
import uuid
import sys
import os
import argparse
import win32com.client

class LogicalError(Exception):
    def __init__(self, message):
        self.message = message

def parseArgs():
    parser = argparse.ArgumentParser(description="Create properties in Renga project")
    parser.add_argument("--project", dest = "project", help="Project path", required=True)
    parser.add_argument("--properties", dest = "properties", help="Properties JSON path", required=True)
    return parser.parse_args()

if __name__ == '__main__':
    try: 
        args = parseArgs()
    except:
        exit(1)

    try:
        app = win32com.client.Dispatch("Renga.Application.1")
        app.Visible = False

        # открытие проекта
        print("Открытие проекта: " + args.project)

        result = app.OpenProject(args.project)
        if result != 0:
            raise LogicalError("Ошибка открытия проекта")

        print("Получение проекта")
        project = app.Project
        if project == None:
           raise LogicalError("Ошибка получения проекта")

        property_mng = project.PropertyManager

        # типы свойств, см. http://help.rengabim.com/api/group__properties.html
        prop_type_dict = {"Double": 1, "String": 2, "Angle": 3, "Area": 4, "Boolean": 5, "Enum": 6, "Int": 7, "Length": 8,
                          "Logical": 9, "Mass": 10, "Volume": 11}  # типы свойств, см. http://help.rengabim.com/api/group__properties.html

        # парсинг json
        with open(args.properties, encoding="utf-8") as properties:
            # загрузка данных
            data = json.load(properties)

            for element, values in data.items():
                # имя секции принимаем за имя свойства
                prop_name = element
                # определяем тип объекта
                prop_type = values['property_type']
                prop_type_id = prop_type_dict.get(prop_type)

                # описание свойства
                prop_desc = property_mng.CreatePropertyDescription(
                    prop_name, prop_type_id)
                # добавление списка, если тип свойства - перечисление
                if prop_type == "Enum":
                    enumeration_items = values['list']
                    prop_desc.SetEnumerationItems(enumeration_items)

                # определение uuid
                if 'id' in values:
                    prop_id = values['id']
                else:
                    prop_id = uuid.uuid1()

                # регистрация свойства
                property_mng.RegisterPropertyS2(prop_id, prop_desc)

                #присваивание свойства всем выбранным типам объектов. Типы объектов задаются по # uuid типа объекта, которому нужно присвоить заданные значения, см. http://help.rengabim.com/api/group___object_types.html
                if 'object_type' in values:
                    for object_type_id in values['object_type']:
                        property_mng.AssignPropertyToTypeS(prop_id, object_type_id)
                    print("Свойство " + prop_name + " назначено")

        print("Сохранение проекта")
        result = project.Save()
        if result != 0:
            raise LogicalError("Ошибка сохранения")

        print("Закрытие проекта")
        result = app.CloseProject(1)
        if result != 0:
            raise LogicalError("Ошибка закрытия проекта")

    except LogicalError as error:
        print(error.message)
        app.Quit()
        exit(1)

    except:
        print("Произошла ошибка")
        app.Quit()
        exit(1)

    print("Закрытие приложения")
    app.Quit()
    exit(0)
