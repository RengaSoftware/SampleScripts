import argparse
import win32com.client


def parseArgs():
    parser = argparse.ArgumentParser(description="Print topic names")
    parser.add_argument("--project", dest="projectPath", help="Project path", required=True)
    parser.add_argument("--topic", dest="topicName",help="Topic name", required=True)
    parser.add_argument("--pdf", dest="pdfPath", help="Pdf path", required=True)                    

    return parser.parse_args()

# функция получения имени раздела по идентификатору раздела
def getTopicName(project, topicId):
    return project.Topics.GetById(topicId).Name

# функция получения списка идентификаторов чертежей в выбранном разделе
def getTopicDrawingIds(project, topicName):
    result = []
    drawingIds = project.Drawings2.GetIds()
    for drawingId in drawingIds:
        # получаем идентификатор чертежа
        drawing = project.Drawings2.GetById(drawingId)
        #получаем параметры чертежа
        params = drawing.GetInterfaceByName('IParameterContainer')
        # из парамеров выбираем раздел, подставляем идентификатор параметра из справки https://help.rengabim.com/api/group__parameter_ids.html
        topicParam = params.GetS("3B7FDF99-6C5E-4FED-8A3C-42149FE5D8B4")
        # получаем идентификатор раздела чертежа 
        topicId = topicParam.GetIntValue()
        # проверяем, что чертежу назначен раздел
        if topicId != 1 :
        # проверям какой раздел назначен чертежу, если раздел подходит, добавляем идентификатор в список
            if getTopicName(project, topicId) == topicName:
                result.append(drawing.UniqueIdS)
    return result

# функция получения идентификаторов чертежей отсортиванных чертежей по номерам

def getSortedDrawingsByNumber(project, drawingIds):
    return sorted(drawingIds, key=lambda id: project.GetEntityNumberInTopicS(id))


if __name__ == '__main__':
    try:
        args = parseArgs()
    except:
        exit(1)

    try:
        app = win32com.client.Dispatch("Renga.Application.1")
        app.Visible = False

        # открытие проекта
        print("Открытие проекта: " + args.projectPath)

        result = app.OpenProject(args.projectPath)
        if result != 0:
            raise Exception("Ошибка открытия проекта")

        print("Получение проекта")
        project = app.Project
        if project == None:
           raise Exception("Ошибка получения проекта")

        # получение идентификаторов чертежей в выбранном разделе
        topicDrawingIds = getTopicDrawingIds(project, args.topicName)

        # сортировка списка чертежей по номерам
        sortedDrawings = getSortedDrawingsByNumber(project, topicDrawingIds)

        # пакетный экспорт в PDF отсортированных чертежей
        print("Идёт экспорт... Не закрывайте окно")
        project.ExportDrawingsToPdfS(sortedDrawings, args.pdfPath, True)
        print("Чертежи экпортированы")

        print("Закрытие проекта")
        result = app.CloseProject(1)
        if result != 0:
            raise Exception("Ошибка закрытия проекта")

    except Exception as error:
        print(error)
        app.Quit()
        exit(1)

    print("Закрытие приложения")
    app.Quit()
    exit(0)
