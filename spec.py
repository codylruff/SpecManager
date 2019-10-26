class spec():
    
    def __init__(self, material_id, process_id, properties):
        self.material_id = material_id
        self.process_id = process_id
        self.properties = properties

    def change_property(property_name, value):
        self.properties[property_name] = value

    def get_properties_json():
        return 