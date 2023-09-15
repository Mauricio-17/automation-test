import pandas as pd
import plotly.graph_objects as go
import os

df = pd.read_excel("./data/result.xlsx", sheet_name='Secondary')
df_t = df.T
fig_t = go.Figure(data=[go.Table(header=dict(values=df.keys()),
                                 cells=dict(values=[i for i in df_t.values])
                                 )])
if __name__ == '__main__':
    if os.path.exists("data"):
        print("Does exist")
        fig_t.write_image("./data/my_own.png")
