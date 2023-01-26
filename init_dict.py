import pickle


def restart_dict():
    p = "C:\\Users\\tigerault\\PycharmProjects\\Multitag\\Storage\\"
    confirm = input("Are you sure you want to (re)initialize blocs and plan? (y/n)")
    if confirm == 'y':

        d = dict(text=['default'], source=['default'], tag=[[]])
        with open(p + "blocs.pkl", 'wb') as f:
            pickle.dump(d, f)

        d = dict(position=[], ID=[], order=[], note=[])
        with open(p + "plan.pkl", 'wb') as f:
            pickle.dump(d, f)


restart_dict()
