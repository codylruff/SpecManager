import json

class Spec():
    
    def __init__(self, material_id, spec_template, properties):
        self.material_id = material_id
        self.spec_template = spec_template
        self.properties = properties

    def change_property(property_name, value):
        self.properties[property_name] = value

    def get_properties_json():
        return json.dumps(self.properties)

    def get_process_id():
        return self.spec_template.process_id

    def get_spec_type():
        return self.spec_template.spec_type