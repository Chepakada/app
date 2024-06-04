import streamlit as st
import pandas as pd
import plotly.graph_objects as go

def main():
    with st.sidebar:
        
        # Create columns for inputs
        col1, col2 = st.columns(2)
        
        # Inputs in the first column
        with col1:
            creators = st.slider("Tableau Creators", min_value=1, max_value=10, value=2, step=1, key="creators")
            explorers = st.slider("Tableau Explorers", min_value=1, max_value=10, value=2, step=1, key="explorers")
            viewers = st.slider("Tableau Viewers", min_value=1, max_value=10, value=2, step=1, key="viewers")
        
        # Inputs in the second column
        with col2:
            ds_salary = st.slider("Data Scientist Salary", min_value=60000, max_value=145000, value=125000, step=5000, key="ds_salary")
            ba_salary = st.slider("Business Analyst Salary", min_value=45000, max_value=145000, value=90000, step=5000, key="ba_salary")

    rcosts = ds_salary
    pcosts = ds_salary
    tcosts = ((75 * creators + 42 * explorers + 15 * viewers) + ba_salary)

    data = pd.DataFrame({
            "Tool": ["Streamlit Cost", "R Shiny Cost", "Tableau Cost"],
            "Cost": [pcosts, rcosts, tcosts]
        })

    # Plot the graph
    st.subheader("Annual Cost for Different Tools")
    fig = go.Figure([go.Bar(x=data["Tool"], y=data["Cost"], marker_color="lightblue")])
    fig.update_layout(xaxis_title="Tool", yaxis_title="Cost")
    st.plotly_chart(fig)

if __name__ == "__main__":
    main()
