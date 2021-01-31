class Person:
    def __init__(self, last_name, first_name, institution, program, position, knowledge, email, website,
                 gender, urm, date_modified, status):
        if first_name is not None and last_name is not None:
            self.name = first_name + ' ' + last_name
        else:
            def select_name(first_name, last_name):
                if first_name is None:
                    return str(last_name)
                elif last_name is None:
                    return str(first_name)
                else:
                    return ""
            self.name = 'INCOMPLETE: ' + select_name(first_name, last_name)
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
        self.all_attributes = [last_name, first_name, institution, program,
                               position, knowledge, email, website, gender,
                               urm, date_modified, status]


class Department:
    def __init__(self, name, people):
        self.name = name
        self.people = people


