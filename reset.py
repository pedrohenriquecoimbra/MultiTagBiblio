import pickle


def restart_dict():
    p = "C:\\Users\\tigerault\\PycharmProjects\\Multitag\\Storage\\"
    d = dict(text=['default'], source=['default'], tag=[[]])
    with open(p + "blocs.pkl", 'wb') as f:
        pickle.dump(d, f)

    d = dict(position=[], ID=[], order=[], note=[])
    with open(p + "plan.pkl", 'wb') as f:
        pickle.dump(d, f)


restart_dict()
