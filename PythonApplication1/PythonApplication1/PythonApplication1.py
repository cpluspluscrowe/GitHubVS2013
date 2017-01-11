def line_profiler(view=None, extra_view=None):
    import line_profiler

    def wrapper(view):
        def wrapped(*args, **kwargs):
            prof = line_profiler.LineProfiler()
            prof.add_function(view)
            if extra_view:
                [prof.add_function(v) for v in extra_view]
            with prof:
                resp = view(*args, **kwargs)
            prof.print_stats()
            return resp
        return wrapped
    if view:
        return wrapper(view)
    return wrapper

@line_profiler
def abc():
    i = 5
    for x in range(10):
        print(x)
    class a():
        def __init__(self):
            self.t = 10
        def add(self,i,j):
            pass
    thing = a()
    thing.add(5,5)

abc()