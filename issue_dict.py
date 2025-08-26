class IssueData:
    def __init__(self):
        self.id          = None
        self.subject     = None
        self.assigned_to = None
        self.start_date  = None
        self.due_date    = None
        self.closed_on   = None
        self.done_ratio  = None

        self.parent_id   = None
        self.children_id = list()

