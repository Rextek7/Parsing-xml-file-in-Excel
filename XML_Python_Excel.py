import pandas as pd
import xml.etree.ElementTree as ET

# Function to parse an XML file and return the root element
def parse_xml(xml_file: str):
    tree = ET.parse(xml_file)
    return tree.getroot()

# Function to create an Excel writer object for a given file name
def create_excel_writer(file_name: str):
    return pd.ExcelWriter(file_name)

# Function to extract data from the ClassLibrary element and return it as a DataFrame
def get_class_library_data(ClassLibrary):
    ClassLibrary_arr = []
    id = ClassLibrary.get('id')
    name = ClassLibrary.get('name')
    description = ClassLibrary.get('description')
    version = ClassLibrary.get('version')
    versionDate = ClassLibrary.get('versionDate')
    contentType = ClassLibrary.get('contentType')
    rows = [id, name, description, version, versionDate, contentType]
    ClassLibrary_arr.append(rows)
    return pd.DataFrame(ClassLibrary_arr,
                        columns=['id', 'name', 'description', 'version', 'versionDate', 'contentType'])

# Function to write a DataFrame to an Excel sheet with a given name
def write_to_excel(df, writer, sheet_name):
    df.to_excel(writer, sheet_name)

# path to get translation
language_xml = '{http://schemas.aveva.com/Generic/Engineering.Information/Class.Library/Model/Basic/2012/01}Languages'

# Function to extract data from the ExtensionNamespaces element and return it as a DataFrame
# At each iteration we access the attributes of the ExtensionNamespace element and store its data in a certain order into an array
def get_extension_namespaces_data(ClassLibrary):
    ExtensionNamespaces = ClassLibrary.find(
        '{http://schemas.aveva.com/Generic/Engineering.Information/Class.Library/Model/Basic/2012/01}ExtensionNamespaces')
    ExtensionNamespace_arr = []
    for i in range(0, len(ExtensionNamespaces)):
        ExtensionNamespace = ExtensionNamespaces[i].attrib
        prefix = ExtensionNamespace.get('prefix')
        name = ExtensionNamespace.get('name')
        type = ExtensionNamespace.get('type')
        uri = ExtensionNamespace.get('uri')
        description = ExtensionNamespace.get('description')
        rows = [prefix, name, type, description, uri]
        ExtensionNamespace_arr.append(rows)
    return pd.DataFrame(ExtensionNamespace_arr,
                        columns=['prefix', 'name', 'type', 'description', 'uri'])

# Function to extract data from the ReferenceData element and return it as multiple DataFrames
# Access 4 ReferenceData items in the function and process each of them separately
def get_reference_data(ClassLibrary):
    ReferenceData = ClassLibrary.find(
        '{http://schemas.aveva.com/Generic/Engineering.Information/Class.Library/Model/Basic/2012/01}ReferenceData')
# 4 ReferenceData elements
    NamingAndNumbering = ReferenceData[0]
    Enumerations = ReferenceData[1]
    UoM = ReferenceData[2]
    Taxonomies = ReferenceData[3]

# Process the 1st element of ReferenceData
    Elements = NamingAndNumbering[0]
    Templates = NamingAndNumbering[1]

    # Elements
    # At each iteration we access one element in Elements and add its values in a specific order to the array
    Elements_arr = []
    for i in range(0, len(Elements)):
        Element = Elements[i].attrib
        id = Element.get('id')
        name = Element.get('name')
        description = Element.get('description')
        mandatory = Element.get('mandatory')
        regEx = Element.get('regEx')
        suffix = Element.get('suffix')
        source = Element.get('source')
        rows = [id, name, description, mandatory,
                regEx, suffix, source]
        Elements_arr.append(rows)

    # datasheet for easy presentation in excel
    Elements_res = pd.DataFrame(Elements_arr,
                                columns=['id', 'name', 'description', 'mandatory', 'regEx', 'suffix', 'source'])

    # Templates
    # Create 2 separate arrays for the 2 tabs

    Templates_arr = []
    Template_Element_arr = []

    # For the 2nd tab we also get id from Template from the 1st tab, so we use a nested loop
    for i in range(0, len(Templates)):
        Template = Templates[i].attrib
        Template_id = Template.get('id')
        name = Template.get('name')
        description = Template.get('description')
        applicableFor = Template.get('applicableFor')
        rows = [Template_id, name, description, applicableFor]
        Templates_arr.append(rows)
        for j in range(0, len(Templates[i][0])):
            Template_Element = Templates[i][0][j].attrib
            id = Template_Element.get('id')
            sortOrder = Template_Element.get('sortOrder')
            description = Template_Element.get('description')
            hideOnEmptyValue = Template_Element.get('hideOnEmptyValue')
            mandatory = Template_Element.get('mandatory')
            prefix = Template_Element.get('prefix')
            regEx = Template_Element.get('regEx')
            Template_Element_rows = [Template_id, id, sortOrder, description, hideOnEmptyValue, mandatory, prefix,
                                     regEx]
            Template_Element_arr.append(Template_Element_rows)

    # datasheets for easy presentation in excel
    Templates_res = pd.DataFrame(Templates_arr, columns=['id', 'name', 'description', 'applicableFor'])
    Template_Element_res = pd.DataFrame(Template_Element_arr,
                                        columns=['Template_id', 'id', 'sortOrder', 'description', 'hideOnEmptyValue',
                                                 'mandatory', 'prefix', 'regEx'])
    # Process the 2nd element of ReferenceData
    # Enumerations
    Enumerations_arr = []

    # In the loop, we access each element in Enumerations and add its values in a certain order to the array
    for i in range(0, len(Enumerations)):
        List = Enumerations[i].attrib
        id = List.get('id')
        sortedOrder = ''
        aspect = List.get('aspect')
        name = List.get('name')
        description = List.get('description')
        # To get its name and description in Russian we use try catch, because not all elements have a translation
        try:
            Languages = Enumerations[i].find(language_xml)
            Language = Languages[0]

        except TypeError:
            Name_ru = ''
            description_ru = ''
            rows = [id, sortedOrder, aspect, name, Name_ru, description, description_ru]
            Enumerations_arr.append(rows)

        else:
            Languages = Enumerations[i].find(language_xml)
            Language = Languages[0]
            Name_ru = Language.get('name')
            description_ru = Language.get('description')
            rows = [id, sortedOrder, aspect, name, Name_ru, description, description_ru]
            Enumerations_arr.append(rows)


        # Each element in Enumerations has its own subarray, so we use a nested loop to iterate over each element from the subarray of the Enumeration element

        # Checking for a subarray
        if len(Enumerations[i]) == 0:
            id_item = ''
            sortedOrder_item = ''
            aspect_item = ''
            name_item = ''
            description_item = ''
            Name_item_ru = ''
            description_item_ru = ''
            rows = [id_item, sortedOrder_item, aspect_item, name_item, Name_item_ru, description_item,
                    description_item_ru]
            Enumerations_arr.append(rows)
        else:
            for j in range(0, len(Enumerations[i])):
                if Enumerations[i][
                    j].tag == '{http://schemas.aveva.com/Generic/Engineering.Information/Class.Library/Model/Basic/2012/01}Items':
                    for k in range(0, len(Enumerations[i][j])):
                        List_item = Enumerations[i][j][k].attrib
                        id_item = id + '_' + List_item.get('id')
                        sortedOrder_item = List_item.get('sortOrder')
                        aspect_item = ''
                        name_item = List_item.get('name')
                        description_item = List_item.get('description')

                        # To get its name and description in Russian we use try catch, because not all elements have a translation
                        try:
                            Languages_item = Enumerations[i][j][k].find(language_xml)
                            Language_item = Languages_item[0]
                        except TypeError:
                            Name_item_ru = ''
                            description_item_ru = ''
                            rows = [id_item, sortedOrder_item, aspect_item, name_item, Name_item_ru,
                                    description_item,
                                    description_item_ru]
                            Enumerations_arr.append(rows)
                        else:
                            Languages_item = Enumerations[i][j][k].find(language_xml)
                            Language_item = Languages_item[0]
                            Name_item_ru = Language_item.get('name')
                            description_item_ru = Language_item.get('description')
                            rows = [id_item, sortedOrder_item, aspect_item, name_item, Name_item_ru,
                                    description_item,
                                    description_item_ru]
                            Enumerations_arr.append(rows)

    # datasheet for easy presentation in excel
    Enumerations_res = pd.DataFrame(Enumerations_arr,
                                    columns=['id', 'sortedOrder', 'aspect', 'name', 'Name_ru', 'description',
                                             'description_ru'])

    # UoM
    # UoM has 2 subarrays, so we access them separately

    # In the loop, we access each element in Units and add its values in a certain order to the array
    Units = UoM[0]
    Units_arr = []
    for i in range(0, len(Units)):
        Unit = Units[i].attrib
        id = Unit.get('id')
        name = Unit.get('name')
        description = Unit.get('description')
        symbol = Unit.get('symbol')

        # To get its name and description in Russian we use try catch, because not all elements have a translation
        try:
            Languages = Units[i].find(language_xml)
            Language = Languages[0]

        except TypeError:
            Name_ru = ''
            description_ru = ''
            rows = [id, name, Name_ru, description, description_ru, symbol]
            Units_arr.append(rows)

        else:
            Languages = Units[i].find(language_xml)
            Language = Languages[0]
            Name_ru = Language.get('name')
            description_ru = Language.get('description')
            rows = [id, name, Name_ru, description, description_ru, symbol]
            Units_arr.append(rows)

    # datasheet for easy presentation in excel
    Units_res = pd.DataFrame(Units_arr, columns=['id', 'name', 'Name_ru', 'description', 'description_ru', 'symbol'])

    # MeasureClasses
    # In the loop, we access each element in MeasureClasses and add its values in a certain order to the array
    MeasureClasses = UoM[1]
    MeasureClasses_arr = []

    # Items in MeasureClasses also have subarrays, so we put subarrays in a separate tab
    MeasureClasses_Units_arr = []

    for i in range(0, len(MeasureClasses)):
        MeasureClass = MeasureClasses[i].attrib
        id = MeasureClass.get('id')
        name = MeasureClass.get('name')
        description = MeasureClass.get('description')
        sortOrder = ''
        # To get its name and description in Russian we use try catch, because not all elements have a translation
        try:
            Languages = MeasureClasses[i].find(language_xml)
            Language = Languages[0]

        except TypeError:
            name_ru = ''
            description_ru = ''
            rows = [id, name, name_ru, sortOrder, description, description_ru]
            MeasureClasses_arr.append(rows)

        else:
            Languages = MeasureClasses[i].find(language_xml)
            Language = Languages[0]
            name_ru = Language.get('name')
            description_ru = Language.get('description')
            rows = [id, name, name_ru, sortOrder, description, description_ru]
            MeasureClasses_arr.append(rows)

        # Processing of MeasureClasses subarrays

        # subarray check
        if len(MeasureClasses[i]) == 0:
            id_Unit = ''
            sortedOrder_Unit = ''
            name_Unit = ''
            description_Unit = ''
            Name_Unit_ru = ''
            description_Unit_ru = ''
            rows = [id, id_Unit, name_Unit, Name_Unit_ru, sortedOrder_Unit, description_Unit, description_Unit_ru]
            MeasureClasses_Units_arr.append(rows)
        else:
            # For each element in the subarray we get its unique values
            for j in range(0, len(MeasureClasses[i])):
                if MeasureClasses[i][
                    j].tag == '{http://schemas.aveva.com/Generic/Engineering.Information/Class.Library/Model/Basic/2012/01}Units':
                    for k in range(0, len(MeasureClasses[i][j])):
                        MeasureClasses_Unit = MeasureClasses[i][j][k].attrib
                        id_unit = MeasureClasses_Unit.get('id')
                        sortedOrder_unit = MeasureClasses_Unit.get('sortOrder')
                        name_Unit = MeasureClasses_Unit.get('id')
                        # Also implemented possibility of correlation of translation from MeasureClasses with translation of value in subarray
                        Name_Unit_ru = ''
                        description_Unit = ''
                        description_Unit_ru = ''
                        rows = [id, id_unit, name_Unit, Name_Unit_ru, sortedOrder_unit, description_Unit,
                                description_Unit_ru]
                        MeasureClasses_Units_arr.append(rows)

    # datasheet for easy presentation in excel
    MeasureClasses_res = pd.DataFrame(MeasureClasses_arr,
                                      columns=['id', 'name', 'Name_ru', 'sortedOrder', 'description', 'description_ru'])

    # datasheet for easy presentation in excel
    MeasureClasses_Units_res = pd.DataFrame(MeasureClasses_Units_arr,
                                            columns=['Class id', 'id', 'name', 'Name_ru', 'sortedOrder', 'description',
                                                     'description_ru'])

    # Taxonomie
    # In the loop, we access each element in Taxonomie and add its values in a certain order to the array
    Taxonomie_arr = []
    for i in range(0, len(Taxonomies)):
        Taxonomie = Taxonomies[i].attrib
        id = Taxonomie.get('id')
        name = Taxonomie.get('name')
        concept = Taxonomie.get('concept')
        Class_Id = ''
        rows = [concept, id, name, Class_Id]
        Taxonomie_arr.append(rows)
        # Each element in Taxonomie has its own subarray
        # To output the values from the subarray along with the name of the element in Taxonomie we use a nested loop
        for j in range(0, len(Taxonomies[i][0])):
            Node = Taxonomies[i][0][j].attrib
            concept = ''
            id_node = id + '_' + Node.get('id')
            name_node = Node.get('name')
            Class_Id = []
            for k in range(0, len(Taxonomies[i][0][j][0])):
                Class = Taxonomies[i][0][j][0][k].attrib
                Class_Id.append(Class.get('id'))
            Class_Id_output = ", ".join(str(element) for element in Class_Id)
            rows = [concept, id_node, name_node, Class_Id_output]
            Taxonomie_arr.append(rows)

    # datasheet for easy presentation in excel
    Taxonomie_res = pd.DataFrame(Taxonomie_arr, columns=['concept', 'id', 'name', 'Class_Id'])

    return Elements_res, Templates_res, Template_Element_res, Enumerations_res, Units_res, MeasureClasses_res, MeasureClasses_Units_res, Taxonomie_res


# Attributes
def get_attributes_data(ClassLibrary):
    Attributes = ClassLibrary.find(
        '{http://schemas.aveva.com/Generic/Engineering.Information/Class.Library/Model/Basic/2012/01}Attributes')

    Attributes_arr = []

    # In the loop, we access each element in Attributes and add its values in a certain order to the array
    for i in range(0, len(Attributes)):
        Attribute = Attributes[i].attrib
        id = Attribute.get('id')
        name = Attribute.get('name')
        description = Attribute.get('description')
        size = Attribute.get('size')
        presence = Attribute.get('presence')
        groupId = Attribute.get('groupId')
        concept = Attribute.get('concept')
        dataType = Attribute.get('dataType')

        # To get its name and description in Russian we use try catch, because not all elements have a translation
        try:
            Languages = Attributes[i].find(language_xml)
            Language = Languages[0]

        except TypeError:
            Name_ru = ''
            description_ru = ''
            rows = [id, name, Name_ru, description, description_ru, size, presence, groupId, concept, dataType]
            Attributes_arr.append(rows)

        else:
            Languages = Attributes[i].find(language_xml)
            Language = Languages[0]
            Name_ru = Language.get('name')
            description_ru = Language.get('description')
            rows = [id, name, Name_ru, description, description_ru, size, presence, groupId, concept, dataType]
            Attributes_arr.append(rows)

    # datasheet for easy presentation in excel
    Attributes_res = pd.DataFrame(Attributes_arr,
                                  columns=['id', 'name', 'Name_ru', 'description', 'description_ru', 'size', 'presence',
                                           'groupId', 'concept', 'dataType'])
    return Attributes_res


# Functionals
def get_functionals_data(ClassLibrary):
    Functionals = ClassLibrary.find(
        '{http://schemas.aveva.com/Generic/Engineering.Information/Class.Library/Model/Basic/2012/01}Functionals')

    # Since each class has a NamingTemplate and Attribute, we create 3 separate tabs for the class, NamingTemplate and Attribute
    Functionals_Class_arr = []
    Functionals_NamingTemplates_arr = []
    Functionals_Attributes_arr = []

    # In the loop, we access each element in Functionals and add its values in a certain order to the array
    for i in range(0, len(Functionals)):
        Class = Functionals[i].attrib
        Class_id = Class.get('id')
        name = Class.get('name')
        description = Class.get('description')
        abstract = Class.get('abstract')
        extends = Class.get('extends')
        type = Class.get('type')

        # To get its name and description in Russian we use try catch, because not all elements have a translation
        try:
            Languages = Functionals[i].find(language_xml)
            Language = Languages[0]

        except TypeError:
            Name_ru = ''
            description_ru = ''
            rows = [Class_id, name, Name_ru, description, description_ru, abstract, extends, type]
            Functionals_Class_arr.append(rows)

        else:
            Languages = Functionals[i].find(language_xml)
            Language = Languages[0]
            Name_ru = Language.get('name')
            description_ru = Language.get('description')
            rows = [Class_id, name, Name_ru, description, description_ru, abstract, extends, type]
            Functionals_Class_arr.append(rows)

        # NamingTemplates
        # Check for NamingTemplates
        try:
            NamingTemplates = Functionals[i].find(
                '{http://schemas.aveva.com/Generic/Engineering.Information/Class.Library/Model/Basic/2012/01}NamingTemplates')
            NamingTemplate = NamingTemplates[0]

        except TypeError:
            NamingTemplate_id = ''
            applicableFor = ''
            NamingTemplate_rows = [Class_id, NamingTemplate_id, applicableFor]
            Functionals_NamingTemplates_arr.append(NamingTemplate_rows)

        else:
            # If available, get its data
            NamingTemplates = Functionals[i].find(
                '{http://schemas.aveva.com/Generic/Engineering.Information/Class.Library/Model/Basic/2012/01}NamingTemplates')
            NamingTemplate = NamingTemplates[0]
            NamingTemplate_id = NamingTemplate.get('id')
            applicableFor = NamingTemplate.get('applicableFor')
            NamingTemplate_rows = [Class_id, NamingTemplate_id, applicableFor]
            Functionals_NamingTemplates_arr.append(NamingTemplate_rows)

        # Functionals_Attributes
        # In the loop, we access each element in Attribute and add its values in a certain order to the array
        # Since we need the class id, we use a nested loop for attribute
        for j in range(0, len(Functionals[i].find(
                '{http://schemas.aveva.com/Generic/Engineering.Information/Class.Library/Model/Basic/2012/01}Attributes')
        )):
            Functionals_Attribute = Functionals[i].find(
                '{http://schemas.aveva.com/Generic/Engineering.Information/Class.Library/Model/Basic/2012/01}Attributes')[
                j].attrib
            Functionals_Attributes_id = Functionals_Attribute.get('id')
            Functionals_Attributes_name = Functionals_Attribute.get('name')
            Functionals_Attributes_description = Functionals_Attribute.get('description')
            Functionals_Attributes_size = Functionals_Attribute.get('size')
            Functionals_Attributes_presence = Functionals_Attribute.get('presence')
            Functionals_Attributes_validationType = Functionals_Attribute.get('validationType')
            Functionals_Attributes_validationRule = Functionals_Attribute.get('validationRule')
            Attributes_rows = [Class_id, Functionals_Attributes_id, Functionals_Attributes_name,
                               Functionals_Attributes_description,
                               Functionals_Attributes_size, Functionals_Attributes_presence,
                               Functionals_Attributes_validationType, Functionals_Attributes_validationRule]
            Functionals_Attributes_arr.append(Attributes_rows)

    # datasheets for easy presentation in excel
    Functionals_Class_res = pd.DataFrame(Functionals_Class_arr,
                                         columns=['Class_id', 'name', 'Name_ru', 'description', 'description_ru',
                                                  'abstract', 'extends', 'type'])

    Functionals_NamingTemplates_res = pd.DataFrame(Functionals_NamingTemplates_arr,
                                                   columns=['Class_id', 'NamingTemplate_id', 'applicableFor'])

    Functionals_Attributes_res = pd.DataFrame(Functionals_Attributes_arr,
                                              columns=['Class_id', 'Attributes_id', 'name', 'description',
                                                       'size', 'presence', 'validationType', 'validationRule'])

    return Functionals_Class_res, Functionals_NamingTemplates_res, Functionals_Attributes_res


# Generals
def get_generals_data(ClassLibrary):
    Generals = ClassLibrary.find(
        '{http://schemas.aveva.com/Generic/Engineering.Information/Class.Library/Model/Basic/2012/01}Generals')

    # Since classes have subarrays, we create 2 arrays for different tabs
    Generals_Class_arr = []
    Generals_Attributes_arr = []

    # In the loop, we access each element in Generals and add its values in a certain order to the array
    for i in range(0, len(Generals)):
        General_Class = Generals[i].attrib
        Class_id = General_Class.get('id')
        obsolete = General_Class.get('obsolete')
        sortOrder = General_Class.get('sortOrder')
        name = General_Class.get('name')
        description = General_Class.get('description')
        abstract = General_Class.get('abstract')
        extends = General_Class.get('extends')

        # To get its name and description in Russian we use try catch, because not all elements have a translation
        try:
            Languages = Generals[i].find(language_xml)
            Language = Languages[0]

        except TypeError:
            Name_ru = ''
            description_ru = ''
            rows = [Class_id, name, Name_ru, description, description_ru, obsolete, sortOrder, abstract, extends]
            Generals_Class_arr.append(rows)

        else:
            Languages = Generals[i].find(language_xml)
            Language = Languages[0]
            Name_ru = Language.get('name')
            description_ru = Language.get('description')
            rows = [Class_id, name, Name_ru, description, description_ru, obsolete, sortOrder, abstract, extends]
            Generals_Class_arr.append(rows)

        # subarray check
        try:
            Attributes = Generals[i].find(
                '{http://schemas.aveva.com/Generic/Engineering.Information/Class.Library/Model/Basic/2012/01}Attributes')
            Attribute = Attributes[0]

        except TypeError:
            Attribute_id = ''
            rows = [Class_id, Attribute_id]
            Generals_Attributes_arr.append(rows)

        # Get the value in the subarray together with the item name
        else:
            Attributes = Generals[i].find(
                '{http://schemas.aveva.com/Generic/Engineering.Information/Class.Library/Model/Basic/2012/01}Attributes')
            Attribute = Attributes[0]
            for j in range(0, len(Attributes)):
                Attribute = Attributes[j]
                Attribute_id = Attribute.get('id')
                rows = [Class_id, Attribute_id]
                Generals_Attributes_arr.append(rows)


    # datasheets for easy presentation in excel
    Generals_Class_res = pd.DataFrame(Generals_Class_arr,
                                      columns=['Class_id', 'name', 'Name_ru', 'description', 'description_ru',
                                               'obsolete',
                                               'sortOrder', 'abstract', 'extends'])

    Generals_Attributes_res = pd.DataFrame(Generals_Attributes_arr, columns=['Class_id', 'id'])

    return Generals_Class_res, Generals_Attributes_res

def write_data_to_excel(writer, ClassLibrary):

    # ClassLibrary
    ClassLibrary_res = get_class_library_data(ClassLibrary)
    write_to_excel(ClassLibrary_res, writer, 'ClassLibrary')

    # ExtensionNamespaces
    ExtensionNamespace_res = get_extension_namespaces_data(ClassLibrary)
    write_to_excel(ExtensionNamespace_res, writer, 'ExtensionNamespaces')

    # ReferenceData
    Elements_res, Templates_res, Template_Element_res, Enumerations_res, Units_res, MeasureClasses_res, MeasureClasses_Units_res, Taxonomie_res = get_reference_data(ClassLibrary)
    write_to_excel(Elements_res, writer, 'N&N Elements')
    write_to_excel(Templates_res, writer, 'N&N Templates')
    write_to_excel(Template_Element_res, writer, 'N&N Template Elements')
    write_to_excel(Enumerations_res, writer, 'Enumerations')
    write_to_excel(Units_res, writer, 'Units')
    write_to_excel(MeasureClasses_res, writer, 'MeasureClasses')
    write_to_excel(MeasureClasses_Units_res, writer, 'MeasureClasses Units')
    write_to_excel(Taxonomie_res, writer, 'Taxonomie')

    # Attributes
    Attributes_res = get_attributes_data(ClassLibrary)
    write_to_excel(Attributes_res, writer, 'Attributes')

    # Functionals
    Functionals_Class_res, Functionals_NamingTemplates_res, Functionals_Attributes_res = get_functionals_data(ClassLibrary)
    write_to_excel(Functionals_Class_res, writer,
                   'Functionals Class')
    write_to_excel(Functionals_NamingTemplates_res,
                   writer, 'Functionals NamingTemplates')
    write_to_excel(Functionals_Attributes_res,
                   writer, 'Functionals Attributes')

    # Generals
    Generals_Class_res, Generals_Attributes_res = get_generals_data(ClassLibrary)
    write_to_excel(Generals_Class_res,
                   writer, 'Generals Class')
    write_to_excel(Generals_Attributes_res,
                   writer, 'Generals Attributes')


if __name__ == "__main__":
    xml_file = "AVEVA ISM Standard Xml Class Library format.xml"
    ClassLibrary = parse_xml(xml_file)

    # create excel writer object
    writer = create_excel_writer('data_3.xlsx')

    # Write data to excel
    write_data_to_excel(writer, ClassLibrary)

    # save the excel
    writer._save()