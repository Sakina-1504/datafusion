class DataStore:
    """
    Central memory storage for whole app
    (acts like backend state manager)
    """

    def __init__(self):
        self.data = {
            "uploads": {},
            "validated": {},
            "consolidated": None
        }

    def add_upload(self, file_name, data):
        self.data["uploads"][file_name] = data

    def get_uploads(self):
        return self.data["uploads"]

    def set_validated(self, file_name, data):
        self.data["validated"][file_name] = data

    def get_validated(self):
        return self.data["validated"]

    def set_consolidated(self, data):
        self.data["consolidated"] = data

    def get_consolidated(self):
        return self.data["consolidated"]