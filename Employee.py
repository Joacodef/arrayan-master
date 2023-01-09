class Employee:
    def __init__(self, name, jobs_List):
        self.name = name
        self.jobs_list = jobs_List  # List that contains dictionaries, each containing all the procedures and respective 
                                    # payments for that person.
                                    # Each dictionary corresponds to a different kind of job that the person did for those procedures.
                                    # (Each job is specified in a different sheet in the input excel)
    