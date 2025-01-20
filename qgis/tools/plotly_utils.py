# -*- coding: utf-8 -*-
"""
plotly_utils
---------
"""
from qgis.PyQt.QtWebKitWidgets import QWebView

import plotly.graph_objects as go
import plotly.offline as po

from . import utils

#---------------------------------------------------------------------------------------------------------    
# 
#---------------------------------------------------------------------------------------------------------    
def initPlotly(oDialog):

    try:

        # region categories
        # Diagrams
        # 1. Create  a Plotly figure
        categories_data = [0, 0, 0]
        categories_bar = go.Figure(
            data=[
                go.Bar(
                    x=[
                        "Bygninger",
                        "Veje og andre hårde<br>overflader",
                        "Landskab og andre<br>grønne overflader",
                    ],
                    y=categories_data,
                    width=0.3,
                    name="Totale emissioner",
                    hovertemplate="%{y} tons CO<sub>2</sub><extra></extra>",
                    marker_color=["#FF0000", "#1E90FF", "#228B22"],
                )
            ],
            layout=go.Layout(
                title_text="Totale emissioner fra overordnede kategorier",
                title_x=0.5,
                height=400,
                margin=dict(l=0, r=0, t=100, b=0),
                yaxis=dict(
                    title="tons CO<sub>2</sub>",
                    tick0=0,
                    dtick=5000,
                    autorange=True,
                    showgrid=True,
                    gridcolor="lightgray",
                ),
                plot_bgcolor="#F8F9F9",
            ),
        )

        # 2. Convert categories_bar to HTML
        categories_html = po.plot(
            categories_bar, output_type="div", include_plotlyjs="cdn"
        )
        # 3. Create a QWebView instance
        categories_web_view = QWebView()
        # 4. Set HTML content
        categories_web_view.setHtml(categories_html)
        # 5. Add the QWebView instance to the layout
        oDialog.tabLayout_1.addWidget(categories_web_view)
        # oDialog.verticalLayout_2.addWidget(web_view)
        # endregion

        # region lokal plan
        lokal_data = [568, 2873]
        lokal_bar = go.Figure(
            data=[
                go.Bar(
                    x=["Kan reguleres", "Kan reguleres indirekte"],
                    y=lokal_data,
                    width=0.5,
                    name="Regulering i en lokalplan",
                    hovertemplate="%{y} tons CO<sub>2</sub><extra></extra>",
                    marker_color=["#B5AF98", "#BFBFBF"],
                )
            ],
            layout=go.Layout(
                title_text="Totale emissioner til regulering i en lokalplan",
                title_x=0.5,
                height=400,
                margin=dict(l=0, r=0, t=100, b=0),
                yaxis=dict(
                    title="tons CO<sub>2</sub>",
                    tick0=0,
                    dtick=500,
                    autorange=True,
                    showgrid=True,
                    gridcolor="lightgray",
                ),
                plot_bgcolor="#F8F9F9",
            ),
        )

        lokal_html = po.plot(lokal_bar, output_type="div", include_plotlyjs="cdn")
        lokal_web_view = QWebView()
        lokal_web_view.setHtml(lokal_html)
        oDialog.tabLayout_2.addWidget(lokal_web_view)

        # endregion

        # region benchmark

        benchmark_data = [11, 6]
        benchmark_bar = go.Figure(
            data=[
                go.Bar(
                    x=[""],
                    y=benchmark_data,
                    width=0.3,
                    name="Gældende benchmark",
                    hovertemplate="%{y} kg CO<sub>2</sub>/m<sup>2</sup>/år<extra></extra>",
                    marker_color=["#668FA6"],
                    showlegend=False,
                ),
                go.Scatter(
                    x=[None],
                    y=[None],
                    mode="lines",
                    name="Reduktioner Roadmap (2025)",
                    line=dict(color="#537C37", width=2, dash="dash"),
                    showlegend=True,
                ),
                go.Scatter(
                    x=[None],
                    y=[None],
                    mode="lines",
                    name="Lavemissionsklassen (gældende)",
                    line=dict(color="#FFCFBB", width=2, dash="dash"),
                    showlegend=True,
                ),
                go.Scatter(
                    x=[None],
                    y=[None],
                    mode="lines",
                    name="Bygningsreglementet (gældende)",
                    line=dict(color="#FF8855", width=2, dash="dash"),
                    showlegend=True,
                ),
            ],
            layout=go.Layout(
                title_text="Resultater sammenlignet med gældende benchmark",
                title_x=0.5,
                height=400,
                margin=dict(l=0, r=0, t=100, b=0),
                yaxis=dict(
                    title="kg CO<sub>2</sub>/m<sup>2</sup>/år",
                    tick0=0,
                    dtick=2,
                    autorange=True,
                    showgrid=True,
                    gridcolor="lightgray",
                ),
                shapes=[
                    dict(
                        type="line",
                        xref="paper",
                        x0=0,
                        x1=1,
                        yref="y",
                        y0=5.9,
                        y1=5.9,
                        line=dict(color="#537C37", width=2, dash="dash"),
                    ),
                    dict(
                        type="line",
                        xref="paper",
                        x0=0,
                        x1=1,
                        yref="y",
                        y0=8,
                        y1=8,
                        line=dict(color="#FFCFBB", width=2, dash="dash"),
                    ),
                    dict(
                        type="line",
                        xref="paper",
                        x0=0,
                        x1=1,
                        yref="y",
                        y0=12,
                        y1=12,
                        line=dict(color="#FF8855", width=2, dash="dash"),
                    ),
                ],
                legend=dict(
                    orientation="h", yanchor="bottom", y=-0.2, xanchor="center", x=0.5
                ),
                plot_bgcolor="#F8F9F9",
            ),
        )

        benchmark_html = po.plot(
            benchmark_bar, output_type="div", include_plotlyjs="cdn"
        )
        benchmark_web_view = QWebView()
        benchmark_web_view.setHtml(benchmark_html)
        oDialog.tabLayout_3.addWidget(benchmark_web_view)

        # endregion


    except Exception as e:
        utils.Error("Fejl i initPlotly: " + str(e))

#---------------------------------------------------------------------------------------------------------    
# 
#---------------------------------------------------------------------------------------------------------    
def updateWidget(oTablayout, sumTotalEmissionsBuildings, sumTotalEmissionsType, sumTotalEmissionsTrees):

    try:

        # 1. Create  a Plotly figure
        categories_data = [sumTotalEmissionsBuildings/1000, sumTotalEmissionsType/1000, sumTotalEmissionsTrees/1000]
        maxNumericValue = abs(sumTotalEmissionsBuildings/1000)
        if abs(sumTotalEmissionsType/1000) > maxNumericValue:
            maxNumericValue = abs(sumTotalEmissionsType/1000)
        if abs(sumTotalEmissionsTrees/1000) > maxNumericValue:
            maxNumericValue = abs(sumTotalEmissionsTrees/1000)
        dtickY = 5000
        if maxNumericValue > 500000:
            dtickY = 100000
        elif maxNumericValue > 300000:
            dtickY = 50000
        elif maxNumericValue > 100000:
            dtickY = 10000

        categories_bar = go.Figure(
            data=[
                go.Bar(
                    x=[
                        "Bygninger",
                        "Veje og andre hårde<br>overflader",
                        "Landskab og andre<br>grønne overflader",
                    ],
                    y=categories_data,
                    width=0.3,
                    name="Totale emissioner",
                    hovertemplate="%{y:.0f} tons CO<sub>2</sub><extra></extra>",
                    marker_color=["#FF0000", "#1E90FF", "#228B22"],
                )
            ],
            layout=go.Layout(
                title_text="Totale emissioner fra overordnede kategorier",
                title_x=0.5,
                height=400,
                margin=dict(l=0, r=0, t=100, b=0),
                yaxis=dict(
                    title="tons CO<sub>2</sub>",
                    tick0=0,
                    dtick=dtickY,
                    autorange=True,
                    showgrid=True,
                    gridcolor="lightgray",
                ),
                plot_bgcolor="#F8F9F9",
            ),
        )

        # 2. Convert categories_bar to HTML
        categories_html = po.plot(
            categories_bar, output_type="div", include_plotlyjs="cdn"
        )
        # 3. Create a QWebView instance
        categories_web_view = QWebView()
        # 4. Set HTML content
        categories_web_view.setHtml(categories_html)
        # 5. Remove/add the QWebView instance to the layout
        for i in reversed(range(oTablayout.count())): 
            oTablayout.itemAt(i).widget().setParent(None) 
        oTablayout.addWidget(categories_web_view)

    except Exception as e:
        utils.Error("Fejl i updateWidget: " + str(e))

#---------------------------------------------------------------------------------------------------------    
# 
#---------------------------------------------------------------------------------------------------------    
