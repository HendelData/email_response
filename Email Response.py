import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import webbrowser
import statsmodels.api as sm

# READ IN DATA
rr = pd.read_excel('C:/Users/marce/Documents/Python/SWD/Email Response/Email.xlsx', sheet_name='Data')

# CREATE LABEL FOR X AXIS AND COUNTER SERIES FOR USE WITH REGRESSION
rr['xlabel'] = rr['Quarter'].str.slice(0, 2) + '<br>' + rr['Quarter'].str.slice(5, 7)
rr['xCounter'] = range(0, 9)

# COMPUTE REGRESSION COEFFICIENTS
model1 = sm.OLS(rr['Response Rate'], sm.add_constant(rr['xCounter']))
results1 = model1.fit()
model2 = sm.OLS(rr['Completion Rate'], sm.add_constant(rr['xCounter']))
results2 = model2.fit()

# CREATE SERIES FOR POINTS ON THE REGRESSION LINE
rr['ResponseReg'] = results1.params.iloc[1] * rr['xCounter'] + results1.params.iloc[0]
rr['CompletionReg'] = results2.params.iloc[1] * rr['xCounter'] + results2.params.iloc[0]

# CREATE GRAPH
figure1 = px.line(rr, x='Quarter', y='Response Rate', height=300, width=900,
                  title='<b> Response Rate </b>',
                  labels={'Quarter': '', 'Response Rate': ''})

# FORMAT GRAPH
figure1.update_layout(xaxis=dict(tickmode='array', tickvals=[0, 1, 2, 3, 4, 5, 6, 7, 8],
                                 ticktext=rr['xlabel'], tick0=1,
                                 tickfont=dict(family='Arial', size=10, color='black'),
                                 showgrid=False,
                                 linecolor='rgb(230,230,230)'),
                      yaxis=dict(dtick=0.02, tickformat='.0%', ticksuffix=' ',
                                 tickfont=dict(family='Arial', size=10, color='black'),
                                 range=[-0.001, 0.0405],
                                 gridcolor='rgb(230,230,230)',
                                 griddash='dash'),
                      plot_bgcolor='white')

# ADD REGRESSION LINE
figure1.add_trace(
    go.Scatter(x=rr['Quarter'], y=rr['ResponseReg'], name="slope", mode='lines',
               line=dict(dash='2px', color='rgb(222,184,135)'), showlegend=False)
)

# CREATE HOVER DATA
figure1.update_traces(hovertemplate='%{y:.1%} <extra></extra>')

# CREATE GRAPH
figure2 = px.line(rr, x='Quarter', y='Completion Rate', height=300, width=900,
                  title='<b> Completion Rate </b>',
                  labels={'Quarter': '', 'Completion Rate': ''})

# FORMAT GRAPH
figure2.update_layout(xaxis=dict(tickmode='array', tickvals=[0, 1, 2, 3, 4, 5, 6, 7, 8],
                                 ticktext=rr['xlabel'], tick0=1,
                                 tickfont=dict(family='Arial', size=10, color='black'),
                                 showgrid=False,
                                 linecolor='rgb(230,230,230)'),
                      yaxis=dict(dtick=0.50, tickformat='.0%', ticksuffix=' ',
                                 tickfont=dict(family='Arial', size=10, color='black'),
                                 range=[-0.025, 1.10],
                                 gridcolor='rgb(230,230,230)',
                                 griddash='dash'),
                      plot_bgcolor='white')

# ADD REGRESSION LINE
figure2.add_trace(
    go.Scatter(x=rr['Quarter'], y=rr['CompletionReg'], name="slope", mode='lines',
               line=dict(dash='2px', color='rgb(222,184,135)'), showlegend=False)
)

# CREATE HOVER DATA
figure2.update_traces(hovertemplate='%{y:.1%} <extra></extra>')


def combine_plotly_figs_to_html(plotly_figs, html_fname, include_plotlyjs='cdn',
                                separator=None, auto_open=False):
    with open(html_fname, 'w') as f:
        f.write(plotly_figs[0].to_html(include_plotlyjs=include_plotlyjs))
        for fig in plotly_figs[1:]:
            if separator:
                f.write(separator)
            f.write(fig.to_html(full_html=False, include_plotlyjs=False))

    if auto_open:
        import pathlib
        import webbrowser
        uri = pathlib.Path(html_fname).absolute().as_uri()
        webbrowser.open(uri)


combine_plotly_figs_to_html([figure1, figure2], 'C:/Users/marce/Documents/Python/SWD/Email Response/Two Plots.html',
                            include_plotlyjs='cdn', separator=None, auto_open=False)

webbrowser.open("C:/Users/marce/Documents/Python/SWD/Email Response/Two Plots.html")
