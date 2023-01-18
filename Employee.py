import Levenshtein

class Employee:
    def __init__(self, name, jobs_List):
        self.name = name
        self.jobs_list = jobs_List
        """ 
        jobs_list is a list of dictionaries. Each dictionary should have a key "job_name" and
        a set of other keys corresponding to each procedure, and the value of each key should be
        a list of two values: the first one is the amount, and the second one is the total price.
        An example of a jobs_list would be:
        jobs_list = [
            {"job_name":"Job 1", "procedure 1":[1, 1200], "procedure 2":[2, 2300]},
            {"job_name":"Job 2", "procedure 1":[4, 4800], "procedure 2":[1, 1150]},
        ]
        """

    def get_name(self):
        return self.name

    def get_jobs_list(self):
        return self.jobs_list

    def add_job(self, job_dict, job_name=""):
        if not isinstance(job_dict, dict):
            raise TypeError("job_dict must be a dictionary")

        # Add job_name to job_dict if it doesn't have it:
        if "job_name" not in job_dict.keys():
            if job_name == "":
                raise ValueError("job_name must be given if the job_dict doesn't have a key 'job_name'")
            else:
                if not isinstance(job_name, str):
                    raise TypeError("job_name must be a string")
                else:
                    job_dict["job_name"] = job_name
        self.jobs_list.append({"job_name":job_dict["job_name"]})

        # Add procedures to the job, using the add_procedure_job method:
        for procedure in job_dict.keys():
            if procedure != "job_name":
                try:
                    self.add_procedure_job(job_dict["job_name"], procedure, job_dict[procedure])
                except Exception as error:
                    raise error

    def get_job(self, job_name):
        for job in self.jobs_list:
            if job["job_name"] == job_name:
                return job

    def find_job_index(self, job_name):
        for i in range(len(self.jobs_list)):
            if self.jobs_list[i]["job_name"] == job_name:
                return i
        return -1

    def get_job_names(self):
        job_names = []
        for job in self.jobs_list:
            job_names.append(job["job_name"])
        return job_names

    def get_procedures_job(self, job_name):
        job = self.get_job(job_name)
        procedures = []
        proc_names = []
        for key in job.keys():
            if key != "job_name":
                procedures.append(job[key])
                proc_names.append(key)
        return procedures, proc_names
    
    def add_procedure_job(self, job_name, procedure_name, procedure_data):
        job_pos = self.find_job_index(job_name)
        if job_pos == -1:
            raise ValueError("The job " + job_name + " doesn't exist")
        if isinstance(procedure_data, list):
            if len(procedure_data) == 2:
                self.jobs_list[job_pos][procedure_name] = [0,0]
                self.jobs_list[job_pos][procedure_name][0] = int(procedure_data[0])
                self.jobs_list[job_pos][procedure_name][1] = int(procedure_data[1])
            else:
                raise ValueError("procedure_data must be a list of two values, but the list given has " + str(len(procedure_data)))   
        else:
            raise TypeError("procedure_data must be a list. The given type was: " + str(type(procedure_data)))