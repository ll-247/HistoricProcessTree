from anytree import Node


class NewNode(Node):
    def __init__(self, *args, **kwargs):
        super(NewNode, self).__init__(*args, **kwargs)
        self.tags = []
        self.text = ""
        self.command_line = ""
        self.time = ""
        self.unknown = False
        self.parent_process_name = ""
        self.elevation = ""
        self.integrity = ""
        self.domain = ""
        self.accountname = ""
        self.parentSha1 = ""

    def setPID(self, pid):
        self.tags.append(str(pid))