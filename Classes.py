class Person:
    def __init__(self, last_name, first_name, institution, program, position, knowledge, email, website,
                 gender, urm, date_modified, status):
        self.name = first_name + ' ' + last_name
        self.institution = institution
        self.program = program
        self.position = position
        self.knowledge = knowledge
        self.email = email
        self.website = website
        self.gender = gender
        self.urm = urm
        self.date_modified = date_modified
        self.status = status


class Department:
    def __init__(self, name, people):
        self.name = name
        self.people = people

