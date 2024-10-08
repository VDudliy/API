import requests
import json


class Simma_application:
    url=""
    login=""
    password=""

    def __new__(cls, *args, **kwargs):
        return super().__new__(cls)

    def __init__(self):
        "Вызова init для X " + str(self)

    def logining(self,url,login,password):
        payload = {"action": "auth", "login": login, "passwd": password}
        payload = json.dumps(payload)
        response = requests.post(url + "/api_auth", data=payload,
                                 headers={"Content-Type": "application/json; charset=utf8", "dataType": "json"})
        sid=response.json()
        return sid

class Simma_attribute:
    # сначала описываем атрибуты класса Мета-Атрибут
    a_id = ''
    a_name = ""
    a_type = ''
    a_type_name = ''
    a_defval = ''
    a_descr = ''
    a_defval_arr = []
    a_class = []

    # Блок стандартных методов присущих всем классам (создание и инициализация)
    def __new__(cls, *args, **kwargs):
        # print("Вызова new для X " + str(cls))
        return super().__new__(cls)

    def __init__(self):
        "Вызова init для X " + str(self)

    # блок методов класса
    def cattr(self, url, mid,
              sid):  # метод позволяет создать новый атрибут. Подразумевается что все поля атрибута уже заполнены и осталось только создать в новой базе, получить его id и сохранить в объекте
        if self.a_name == '' or self.a_type == '':
            print("Не задано имя или тип нового атрибута. Создание атрибута не возможно")
            return (1)
        payload = {"action": "cattr", "inputAttrName": self.a_name, "inputAttrType": self.a_type,
                   "inputDefVal": self.a_defval,
                   "inputDefValTA": "", "inputDefValArr": self.a_defval_arr, "inputAttrId": "null",
                   "inputAttrDescr": self.a_descr}
        payload = json.dumps(payload)
        response = requests.post(url + "/api_json?mid=" + mid, data=payload,
                                 headers={"Content-Type": "application/json; charset=utf8",
                                          'Cookie': 'sid=' + sid}
                                 )
        new_id = response.json()
        new_id = new_id['state'][2:]
        if new_id[0] != 'a':
            print("Не корректно прошло создание нового атрибута", response.text)
            return (1)
        return new_id

    def ga(self, url, mid, sid):  # метод позволяет получить информацию об атрибуте по его id
        response = requests.post(url + "/api?mid=" + mid + "&action=ga&id=" + self.a_id,
                                 headers={"Content-Type": "application/json; charset=utf8",
                                          'Cookie': 'sid=' + sid}
                                 )
        info_element = json.loads(response.text)
        print(info_element)
        self.a_name = info_element[0][1]
        self.a_descr = info_element[0][2]
        self.a_type_name = info_element[0][3]
        self.a_defval = info_element[0][4]
        self.a_defval_arr = info_element[0][5]
        self.a_type = info_element[0][6]
        self.a_class = info_element[0][8]

        return (0)

    def da(self, url, mid, sid):  # метод позволяет удалить атрибут
        response = requests.post(url + "/api?mid=" + mid + "&action=da&id=" + self.a_id,
                                 headers={"Content-Type": "application/json; charset=utf8",
                                          'Cookie': 'sid=' + sid}
                                 )
        stat_id = response.json()
        stat_id = stat_id['state']
        if stat_id != 0:
            print("Не корректно прошло удаление атрибута", response.text)
            return (1)
        return (stat_id)


class Simma_model:
    # сначала описываем атрибуты класса Модель
    m_id = ''
    m_name = ''
    m_descr = ''
    m_author = ''
    m_date_model = ''
    m_class_count = 0
    m_attribute_count = 0
    m_element_count = 0
    m_meta_releation_count = 0
    m_relation_count = 0
    m_diagram_count = 0
    m_life = ''

    # Блок стандартных методов присущих всем классам (создание и инициализация)
    def __new__(cls, *args, **kwargs):
        # print("Вызова new для X " + str(cls))
        return super().__new__(cls)

    def __init__(self):
        "Вызова init для X " + str(self)

    # блок методов класса
    def gm(self, url, mid, sid):  # Метод позволяет получить и заполнить все атрибуты класса Модель.
        response = requests.get(url + '/api?action=gm&id=' + mid + '&sid=' + sid)
        about_model = response.json()
        self.m_name = about_model[0][1]
        self.m_descr = about_model[0][2]
        self.m_author = about_model[0][3]
        self.m_date_model = about_model[0][4]
        self.m_life = True

    def gmi(self, url, mid, sid):  # Метод позволяет получить расширенную информацию о Модели
        response = requests.get(url + '/api?action=gmi&id=' + mid + '&sid=' + sid)
        about_model = response.json()
        self.m_class_count = about_model[0][1]
        self.m_attribute_count = about_model[1][1]
        self.m_meta_releation_count = about_model[2][1]
        self.m_diagram_count = about_model[3][1]
        self.m_element_count = about_model[4][1]
        self.m_relation_count = about_model[5][1]

    def dmdl(self, url, mid, sid):  # Метод позволяет удалить модель
        payload = {"action": "dmdl", "id": mid}
        payload = json.dumps(payload)
        response = requests.post(url + "api_json?", data=payload,
                                 headers={"Content-Type": "application/json; charset=utf8",
                                          'Cookie': 'sid=' + sid}
                                 )
        stat_id = response.json()
        stat_id = stat_id['state']
        if stat_id != 0:
            print("Не корректно прошло удаление", response.text)
            return (1)
        self.m_life = False
        return (stat_id)

    def gccnt(self, url, mid,
              sid):  # Метод позволяет получить информацию о количестве элементов в каждом классе выбранной модели
        response = requests.post(url + '/api?action=gccnt&mid=' + mid + '&sid=' + sid)
        about_model = response.json()
        return (about_model)


class Simma_class:
    # сначала описываем атрибуты класса Класс
    o_id = ''
    o_name = ''
    o_descr = ''
    o_visible = ''
    o_element_count = 0
    o_element = []
    o_life = ''

    # Блок стандартных методов присущих всем классам (создание и инициализация)
    def __new__(cls, *args, **kwargs):
        # print("Вызова new для X " + str(cls))
        return super().__new__(cls)

    def __init__(self):
        "Вызова init для X " + str(self)

    # блок методов класса
    def gc(self, url, mid, sid, id=True):  # Метод позволяет получить список классов модели.
        if id:
            response = requests.get(url + '/api?action=gc&mid=' + mid + '&sid=' + sid)
        else:
            response = requests.get(url + '/api?action=gca&mid=' + mid + '&id=' + self.o_id + '&sid=' + sid)
        about_class = response.json()
        return (about_class)

    def cc(self, url, mid, sid):  # Метод позволяет создать новый  Класс.
        payload = {'action': 'cc', "mid": mid, 'name': self.o_name, 'descr': self.o_descr,
                   'visible': self.o_visible}
        payload = json.dumps(payload)
        response = requests.post(url + "/api_json?", data=payload,
                                 headers={"Content-Type": "application/json; charset=utf8",
                                          'Cookie': 'sid=' + sid})
        new_id = response.json()
        new_id = new_id['state'][2:]
        if new_id[0] != 'o':
            print("Не корректно прошло создание нового класса", response.text)
            return (1)
        return new_id

    def gccnt(self, url, mid, sid):  # Метод позволяет получить информацию о количестве элементов в классе
        response = requests.post(url + "/api?action=gccnt&mid=" + mid + '&id=' + self.o_id,
                                 headers={"Content-Type": "application/json; charset=utf8",
                                          'Cookie': 'sid=' + sid})
        about_class = response.json()
        if not isinstance(about_class, list):
            return (1)
        return about_class

    def sc(self, url, mid, sid, c_data1):  # Метод позволяет создать новый  Класс.
        payload = {'action': 'sc', "id": self.o_id, 'data1': c_data1}
        payload = json.dumps(payload)
        response = requests.post(url + "/api_json?", data=payload,
                                 headers={"Content-Type": "application/json; charset=utf8",
                                          'Cookie': 'sid=' + sid})
        new_id = response.json()
        new_id = new_id['state']
        if new_id != '0':
            print("Не корректно прошла привязка", response.text)
            return (1)
        return new_id

    def gca(self, url, mid, sid):  # Метод позволяет получить  все атрибуты привязанные к классу.
        response = requests.get(url + '/api?action=gca&mid=' + mid + '&id=' + self.o_id + '&sid=' + sid)
        about_class = response.json()
        return (about_class)

    def sevt(self, url, mid, sid, list_cols, filter):
        payload = {'action': 'sevt', 'id': self.o_id, 'cols': list_cols, 'filters': filter}
        payload = json.dumps(payload)
        response = requests.post(url + '/api_json?', data=payload,
                                 headers={'Content-Type': 'application/json; charset=utf8',
                                          'Cookie': 'sid=' + sid + '; mid=' + mid}
                                 )
        new_id = response.json()
        new_id = new_id['state']
        if new_id != '0':
            print("Не корректно прошла становка последовательности возврата элементов", response.text)
            return (1)
        return new_id

    def gtvt(self, url, mid, sid):
        payload = {'action': 'gtvt', 'id': self.o_id}
        payload = json.dumps(payload)
        response = requests.post(url + '/api_json?', data=payload,
                                 headers={'Content-Type': 'application/json; charset=utf8',
                                          'Cookie': 'sid=' + sid + '; mid=' + mid}
                                 )
        new_id = json.loads(response.text)
        new_id = new_id['data']
        return new_id


class Simma_element:
    # сначала описываем атрибуты класса элемент
    e_id = ""
    e_name = ""
    e_descr = ""
    e_class = ""

    # Блок стандартных методов присущих всем классам (создание и инициализация)
    def __new__(cls, *args, **kwargs):
        # print("Вызова new для X " + str(cls))
        return super().__new__(cls)

    def __init__(self):
        "Вызова init для X " + str(self)

    # блок методов класса
    def ce(self, url, mid, sid, e_class, list_attrib, list_value):  # Метод позволяет создать новый  Элемент.
        list_for_load = []
        for i in list_attrib:
            list_for_load.append([i, list_value[list_attrib.index(i)]])
        payload = {"action": "ce", "id": e_class,
                   "data1": list_for_load}
        payload = json.dumps(payload)
        response = requests.post(url + "/api_json?", data=payload,
                                 headers={"Content-Type": "application/json; charset=utf8",
                                          'Cookie': 'sid=' + sid}
                                 )
        new_id = response.json()
        new_id = new_id['state']
        if new_id[2] != 'e':
            print("Не корректно прошло создание элемента", response.text)
            return (1)
        return new_id[2:]

    def ee(self, url, mid, sid, class_id, list_attrib, list_value):  # Метод позволяет редактировать  Элемент.
        list_for_load = []
        for i in list_attrib:
            list_for_load.append([i, list_value[list_attrib.index(i)]])
        payload = {'action': 'ee', 'id': self.e_id, 'class_id': class_id,
                   'data1': list_for_load}
        payload = json.dumps(payload)
        response = requests.post(url + '/api_json?', data=payload,
                                 headers={'Content-Type': 'application/json; charset=utf8',
                                          'Cookie': 'sid=' + sid + '; mid=' + mid}
                                 )
        new_id = response.json()
        new_id = new_id['state']
        if new_id[2] != 'e':
            print("Не корректно прошло редактирование элемента", response.text)
            return (1)
        return new_id[2:]

    def de(self, url, mid, sid):  # Метод позволяет удалить Элемент.
        payload = {"action": "de", "id": self.e_id}
        payload = json.dumps(payload)
        response = requests.post(url + "/api?action=de&id=" + self.e_id,
                                 headers={"Content-Type": "application/json; charset=utf8",
                                          'Cookie': 'sid=' + sid + '; mid=' + mid}
                                 )
        new_id = response.json()
        print(new_id)
        new_id = new_id['state']
        if new_id != '0':
            print("Не корректно прошло удаление элемента", response.text)
            return (1)
        return new_id[2:]

    def cce(self, url, mid, sid, conn_id, elem_src_id, elem_dst_id, src_attr_id, dst_attr_id,
            type_id):  # Метод позволяет создать новую связь Элементов.
        payload = {"action": "cce", "conn_id": conn_id, "elem_src_id": elem_src_id, "elem_dst_id": elem_dst_id,
                   "src_attr_id": src_attr_id, "dst_attr_id": dst_attr_id, "type_id": type_id}
        payload = json.dumps(payload)
        response = requests.post(url + "/api_json?", data=payload,
                                 headers={"Content-Type": "application/json; charset=utf8",
                                          'Cookie': 'sid=' + sid + '; mid=' + mid}
                                 )
        new_id = response.json()
        print(new_id)
        new_id = new_id['state']
        if new_id[2] != 'r':
            print("Не корректно создалась связь", response.text)
            return (1)
        return new_id[2:]
    #  получить полную информацию по элементу
    def gea(self, url, mid, sid):
        response = requests.get(url + '/api?mid=' + mid + '&action=gea&id=' + self.e_id + '&sid=' + sid)
        about_element = response.json()
        return (about_element)
    # позволяет получить историю изменений элемента по его id
    def geh(self, url, mid, sid):
        response = requests.get(url + '/api?mid=' + mid + '&action=gea&id=' + self.e_id + '&sid=' + sid)
        about_element = response.json()
        return (about_element)


class Simma_Grafics:
    # Блок стандартных методов присущих всем классам (создание и инициализация)
    def __new__(cls, *args, **kwargs):
        # print("Вызова new для X " + str(cls))
        return super().__new__(cls)

    def __init__(self):
        "Вызова init для X " + str(self)
    # Создает новую диаграмму в Симма. На вход подается словарь с именем и описание диагруммы вида {'name': Diagramm_name, 'descr': Diagramm_name}
    def cg(self,url,mid,sid,dict_diagram):
        payload = {"action": "cdiag", "mid": mid, "data1": dict_diagram}
        payload = json.dumps(payload)
        response = requests.post(url + "/api_json", data=payload,
                                 headers={"Content-Type": "application/json; charset=utf8",
                                          'Cookie': 'sid=' + sid}
                                 )
        for_ret = response.json()
        for_ret = for_ret['state']
        simma_id_diagramm = for_ret[2:]
        return (simma_id_diagramm)
    # Выносит элемент на диаграмму. На вход нужно подать - id элемента в СИММа, id диаграммы на которую нужно вынести элемент
    # координаты x,y,z  элемента , id стенсила

    def cdq(self,url,mid,sid,id_element,id_diagram,x,y,z,id_stensil,height,width,name_height,name_width):
        payload = {"action": "cdq", "id":id_element, "data1": id_diagram,
                   "coords": "{\"x\":" + str(x) + ",\"y\":" + str(y)+ ",\"z\":" + str(z) +",\"stens_id\":\"" +id_stensil + "\",\"height\":" + str(height) + ",\"width\":" + str(width) + ",\"name_height\":" + str(name_height) + ",\"name_width\":" + str(name_width)  + "}",
                   "attrs": ["a", "b", "c", "d", "e", "f"]}
        payload = json.dumps(payload)
        response = requests.post(url + "/api_json?mid=" + mid, data=payload,
                                 headers={"Content-Type": "application/json; charset=utf8",
                                          'Cookie': 'sid=' + sid}
                                 )
        print(response.json())