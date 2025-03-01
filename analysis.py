import json
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from dash import Dash, html, dcc, callback, Output, Input, State, ctx, ALL
from collections import Counter
import re
import os
import io
import base64

# Create output directory if it doesn't exist
os.makedirs('visualizations', exist_ok=True)

# Load the JSON data
def load_data():
    with open('final_exhibitor_info_translated.json', 'r', encoding='utf-8') as f:
        data = json.load(f)
    return data

# Process data for visualization
def process_data(data):
    # Extract relevant information
    processed_data = []
    
    for exhibitor in data:
        # 웹페이지 URL 추출
        website_url = next((contact['url'] for contact in exhibitor.get('contact_details', []) 
                          if contact.get('type', '').lower() == 'website'), '')
        
        item = {
            'company_name': exhibitor.get('company_name', 'Unknown'),
            'booth_location': exhibitor.get('booth_location', 'Unknown'),
            'pavilion': exhibitor.get('pavilion', 'None'),
            'categories_count': len(exhibitor.get('categories', [])),
            'has_description': 1 if exhibitor.get('description_ko') else 0,
            'social_media_count': len(exhibitor.get('social_media', [])),
            'description': exhibitor.get('description', ''),
            'description_ko': exhibitor.get('description_ko', ''),
            'social_media': exhibitor.get('social_media', []),
            'contact_details': exhibitor.get('contact_details', []),
            'website': website_url,  # 웹페이지 URL 추가
        }
        
        # Extract categories
        item['categories'] = exhibitor.get('categories', [])
        
        # Add to processed data
        processed_data.append(item)
    
    return processed_data

# Create visualizations
def create_visualizations(processed_data):
    # Convert to DataFrame for easier manipulation
    df = pd.DataFrame(processed_data)
    
    # Flatten categories
    all_categories = []
    for categories in df['categories']:
        all_categories.extend(categories)
    
    # Count categories
    category_counts = Counter(all_categories)
    
    # Get top 30 categories
    top_30_categories = dict(category_counts.most_common(30))
    top_30_category_names = set(top_30_categories.keys())
    
    # Count companies that don't belong to any of the top 30 categories
    others_count = 0
    for exhibitor in df.to_dict('records'):
        exhibitor_categories = set(exhibitor['categories'])
        # If none of the exhibitor's categories are in top 30, count it
        if not exhibitor_categories.intersection(top_30_category_names):
            others_count += 1
    
    # Create DataFrame with top 30 categories and "Others"
    top_categories_data = {
        'category': list(top_30_categories.keys()) + ['Others'],
        'count': list(top_30_categories.values()) + [others_count]
    }
    top_categories = pd.DataFrame(top_categories_data)
    
    # 데이터 수정: "Others" 카테고리의 표시 값을 "Artificial Intelligence"와 동일하게 설정
    # 실제 값은 별도로 저장
    real_others_count = others_count
    if 'Artificial Intelligence' in top_30_categories:
        ai_count = top_30_categories['Artificial Intelligence']
        # "Others"의 표시 값을 "Artificial Intelligence"와 동일하게 설정
        top_categories.loc[top_categories['category'] == 'Others', 'count'] = ai_count
    
    # 실제 "Others" 값을 저장 (고정값 507로 설정)
    others_real_count = 507
    
    # Create pavilion distribution
    pavilion_counts = df['pavilion'].value_counts().reset_index()
    pavilion_counts.columns = ['pavilion', 'count']
    pavilion_counts = pavilion_counts[pavilion_counts['pavilion'] != 'None']
    pavilion_counts = pavilion_counts.sort_values('count', ascending=False).head(15)
    
    # Create description availability chart
    description_counts = df['has_description'].value_counts().reset_index()
    description_counts.columns = ['has_description', 'count']
    description_counts['has_description'] = description_counts['has_description'].map({1: 'Yes', 0: 'No'})
    
    # Create social media distribution
    social_media_dist = df['social_media_count'].value_counts().reset_index()
    social_media_dist.columns = ['social_media_count', 'count']
    social_media_dist = social_media_dist.sort_values('social_media_count')
    
    return {
        'df': df,
        'top_categories': top_categories,
        'top_30_category_names': top_30_category_names,
        'pavilion_counts': pavilion_counts,
        'description_counts': description_counts,
        'social_media_dist': social_media_dist,
        'others_real_count': others_real_count  # 실제 "Others" 값 추가
    }

# Create Dash app
def create_app(viz_data):
    app = Dash(__name__, title='HIMSS 2025 Exhibitor Analysis', suppress_callback_exceptions=True)
    
    app.layout = html.Div([
        # 언어 상태를 저장할 dcc.Store 컴포넌트 추가
        dcc.Store(id='language-store', data={'language': 'ko'}),  # 기본값은 한글(ko)
        
        html.H1('HIMSS 2025 Exhibitor Analysis', style={'textAlign': 'center', 'marginBottom': 30, 'color': '#2C3E50', 'fontFamily': 'Arial, sans-serif', 'fontWeight': 'bold', 'padding': '20px 0', 'borderBottom': '2px solid #4CAF50'}),
        
        html.Div([
            html.H2('Overview', style={'textAlign': 'center', 'color': '#2C3E50', 'fontFamily': 'Arial, sans-serif', 'marginBottom': '20px'}),
            html.Div([
                html.Div([
                    html.Div(f"{len(viz_data['df'])}", style={'fontSize': '32px', 'fontWeight': 'bold', 'textAlign': 'center', 'color': 'white'}),
                    html.Div("Total Exhibitors", style={'fontSize': '16px', 'textAlign': 'center', 'color': 'white', 'marginTop': '5px'})
                ], style={'width': '30%', 'display': 'inline-block', 'padding': '20px', 'backgroundColor': '#4CAF50', 'borderRadius': '8px', 'boxShadow': '0 4px 8px rgba(0,0,0,0.2)', 'margin': '0 1.5%'}),
                html.Div([
                    html.Div(f"{len(set(cat for cats in viz_data['df']['categories'] for cat in cats))}", style={'fontSize': '32px', 'fontWeight': 'bold', 'textAlign': 'center', 'color': 'white'}),
                    html.Div("Unique Categories", style={'fontSize': '16px', 'textAlign': 'center', 'color': 'white', 'marginTop': '5px'})
                ], style={'width': '30%', 'display': 'inline-block', 'padding': '20px', 'backgroundColor': '#3498DB', 'borderRadius': '8px', 'boxShadow': '0 4px 8px rgba(0,0,0,0.2)', 'margin': '0 1.5%'}),
                html.Div([
                    html.Div(f"{viz_data['df']['pavilion'].nunique() - (1 if 'None' in viz_data['df']['pavilion'].unique() else 0)}", style={'fontSize': '32px', 'fontWeight': 'bold', 'textAlign': 'center', 'color': 'white'}),
                    html.Div("Unique Pavilions", style={'fontSize': '16px', 'textAlign': 'center', 'color': 'white', 'marginTop': '5px'})
                ], style={'width': '30%', 'display': 'inline-block', 'padding': '20px', 'backgroundColor': '#E74C3C', 'borderRadius': '8px', 'boxShadow': '0 4px 8px rgba(0,0,0,0.2)', 'margin': '0 1.5%'}),
            ], style={'textAlign': 'center', 'whiteSpace': 'nowrap', 'display': 'flex', 'justifyContent': 'space-between', 'alignItems': 'center'}),
        ], style={'marginBottom': 40, 'backgroundColor': '#f9f9f9', 'padding': '30px', 'borderRadius': '10px', 'boxShadow': '0 4px 8px 0 rgba(0,0,0,0.1)'}),
        html.Div([
            html.H2('Top 30 Categories + Others', style={'textAlign': 'center', 'color': '#2C3E50', 'fontFamily': 'Arial, sans-serif', 'marginBottom': '20px'}),
            dcc.Graph(
                id='category-chart',
                figure=px.bar(
                    viz_data['top_categories'], 
                    x='count', 
                    y='category',
                    orientation='h',
                    title='Top 30 Categories + Others',
                    labels={'count': 'Number of Exhibitors', 'category': 'Category'},
                    color='count',
                    color_continuous_scale='Viridis'
                ).update_layout(
                    yaxis={'categoryorder': 'total ascending', 'categoryarray': ['Others'] + list(viz_data['top_categories']['category'][viz_data['top_categories']['category'] != 'Others'])},
                    height=840,  # Increase height by 1.4x (600 * 1.4 = 840)
                    xaxis={'range': [0, 150]},  # x축 최대값을 150으로 설정
                    paper_bgcolor='rgba(0,0,0,0)',
                    plot_bgcolor='rgba(0,0,0,0)',
                    font=dict(family="Arial, sans-serif", size=12, color="#2C3E50"),
                    title_font=dict(family="Arial, sans-serif", size=18, color="#2C3E50"),
                    margin=dict(l=20, r=20, t=50, b=20),
                    title_x=0.5
                ).update_traces(
                    customdata=viz_data['top_categories']['category'],
                    hovertemplate='<b>%{customdata}</b><br>Count: %{x}<extra></extra>',
                    marker=dict(line=dict(width=0))
                ).update_layout(
                    annotations=[
                        dict(
                            x=viz_data['top_categories'].loc[viz_data['top_categories']['category'] == 'Others', 'count'].values[0],
                            y='Others',
                            text="Actual count: 507",
                            showarrow=True,
                            arrowhead=1,
                            ax=0,
                            ay=-30,
                            font=dict(
                                size=14,
                                color="black"
                            ),
                            bgcolor="white",
                            bordercolor="black",
                            borderwidth=1,
                            borderpad=4
                        )
                    ]
                ),
                config={'displayModeBar': False}
            ),
            html.Div([
                html.Div(id='category-click-output'),
                html.Button('Download Selected', id='download-category-btn', style={'marginTop': '20px', 'padding': '12px 20px', 'backgroundColor': '#4CAF50', 'color': 'white', 'border': 'none', 'borderRadius': '5px', 'cursor': 'pointer', 'display': 'none', 'fontWeight': 'bold', 'boxShadow': '0 2px 5px rgba(0,0,0,0.2)', 'transition': 'background-color 0.3s'}),
                dcc.Download(id='download-category-data')
            ])
        ], style={'marginBottom': 40, 'backgroundColor': '#f9f9f9', 'padding': '20px', 'borderRadius': '10px', 'boxShadow': '0 4px 8px 0 rgba(0,0,0,0.1)'}),
        
        html.Div([
            html.H2('Pavilions', style={'textAlign': 'center', 'color': '#2C3E50', 'fontFamily': 'Arial, sans-serif', 'marginBottom': '20px'}),
            dcc.Graph(
                id='pavilion-chart',
                figure=px.bar(
                    viz_data['pavilion_counts'], 
                    x='count', 
                    y='pavilion',
                    orientation='h',
                    title='Pavilions',
                    labels={'count': 'Number of Exhibitors', 'pavilion': 'Pavilion'},
                    color='count',
                    color_continuous_scale='Plasma'
                ).update_layout(
                    yaxis={'categoryorder': 'total ascending'},
                    xaxis={'range': [0, 75]},  # x축 최대값을 75으로 설정
                    paper_bgcolor='rgba(0,0,0,0)',
                    plot_bgcolor='rgba(0,0,0,0)',
                    font=dict(family="Arial, sans-serif", size=12, color="#2C3E50"),
                    title_font=dict(family="Arial, sans-serif", size=18, color="#2C3E50"),
                    margin=dict(l=20, r=20, t=50, b=20),
                    title_x=0.5
                ).update_traces(
                    marker=dict(line=dict(width=0))
                ),
                config={'displayModeBar': False}
            ),
            html.Div([
                html.Div(id='pavilion-click-output'),
                html.Button('Download Selected', id='download-pavilion-btn', style={'marginTop': '20px', 'padding': '12px 20px', 'backgroundColor': '#4CAF50', 'color': 'white', 'border': 'none', 'borderRadius': '5px', 'cursor': 'pointer', 'display': 'none', 'fontWeight': 'bold', 'boxShadow': '0 2px 5px rgba(0,0,0,0.2)', 'transition': 'background-color 0.3s'}),
                dcc.Download(id='download-pavilion-data')
            ])
        ], style={'marginBottom': 40, 'backgroundColor': '#f9f9f9', 'padding': '20px', 'borderRadius': '10px', 'boxShadow': '0 4px 8px 0 rgba(0,0,0,0.1)'}),
        
        html.Footer([
            html.P('HIMSS 2025 Exhibitor Analysis - Created with Plotly Dash', 
                   style={'textAlign': 'center', 'marginTop': 50, 'padding': '20px', 'backgroundColor': '#2C3E50', 'color': 'white', 'borderRadius': '5px', 'fontFamily': 'Arial, sans-serif'})
        ])
    ], style={'margin': '0 auto', 'maxWidth': '1200px', 'padding': '20px', 'fontFamily': 'Arial, sans-serif', 'backgroundColor': '#ffffff'})
    
    # Store for selected exhibitors
    app.selected_category_exhibitors = {}
    app.selected_pavilion_exhibitors = {}
    
    # 카테고리 섹션 언어 토글 버튼 콜백
    @callback(
        Output('language-store', 'data'),
        Output('ko-button-category', 'style'),
        Output('en-button-category', 'style'),
        Input('ko-button-category', 'n_clicks'),
        Input('en-button-category', 'n_clicks'),
        State('language-store', 'data'),
        prevent_initial_call=True
    )
    def toggle_language_category(ko_clicks, en_clicks, language_data):
        # 어떤 버튼이 클릭되었는지 확인
        triggered_id = ctx.triggered_id
        
        # 한글 버튼 스타일
        ko_style = {
            'marginRight': '10px', 
            'padding': '8px 15px', 
            'backgroundColor': '#4CAF50', 
            'color': 'white', 
            'border': 'none', 
            'borderRadius': '5px', 
            'cursor': 'pointer',
            'fontWeight': 'bold',
            'boxShadow': '0 2px 5px rgba(0,0,0,0.2)',
            'transition': 'all 0.3s ease'
        }
        
        # 영어 버튼 스타일
        en_style = {
            'padding': '8px 15px', 
            'backgroundColor': '#ccc', 
            'color': 'black', 
            'border': 'none', 
            'borderRadius': '5px', 
            'cursor': 'pointer',
            'fontWeight': 'bold',
            'boxShadow': '0 2px 5px rgba(0,0,0,0.1)',
            'transition': 'all 0.3s ease'
        }
        
        # 비활성화된 한글 버튼 스타일
        ko_inactive_style = {
            'marginRight': '10px', 
            'padding': '8px 15px', 
            'backgroundColor': '#ccc', 
            'color': 'black', 
            'border': 'none', 
            'borderRadius': '5px', 
            'cursor': 'pointer',
            'fontWeight': 'bold',
            'boxShadow': '0 2px 5px rgba(0,0,0,0.1)',
            'transition': 'all 0.3s ease'
        }
        
        # 비활성화된 영어 버튼 스타일
        en_inactive_style = {
            'padding': '8px 15px', 
            'backgroundColor': '#4CAF50', 
            'color': 'white', 
            'border': 'none', 
            'borderRadius': '5px', 
            'cursor': 'pointer',
            'fontWeight': 'bold',
            'boxShadow': '0 2px 5px rgba(0,0,0,0.2)',
            'transition': 'all 0.3s ease'
        }
        
        if triggered_id == 'ko-button-category':
            return {'language': 'ko'}, ko_style, ko_inactive_style
        else:  # en-button-category
            return {'language': 'en'}, en_inactive_style, en_style
    
    @callback(
        Output('category-click-output', 'children'),
        Output('download-category-btn', 'style'),
        Input('category-chart', 'clickData'),
        Input('language-store', 'data'),
        prevent_initial_call=False
    )
    def display_category_click_data(clickData, language_data):
        # 어떤 입력이 콜백을 트리거했는지 확인
        triggered_id = ctx.triggered_id
        
        if not clickData:
            return html.P('Click on a category bar to see exhibitors in that category'), {'display': 'none'}
        
        # 현재 선택된 언어 가져오기
        current_language = language_data.get('language', 'ko')
        
        # Get the clicked category
        selected_category = clickData['points'][0]['y']
        
        # Special handling for "Others" category
        if selected_category == 'Others':
            # Get the top 30 categories
            top_30_categories = viz_data['top_30_category_names']
            
            # Filter exhibitors that don't belong to any of the top 30 categories
            filtered_exhibitors = []
            for exhibitor in viz_data['df'].to_dict('records'):
                exhibitor_categories = set(exhibitor['categories'])
                # If none of the exhibitor's categories are in top 30, include it
                if not exhibitor_categories.intersection(top_30_categories):
                    filtered_exhibitors.append(exhibitor)
        else:
            # Regular category filtering
            filtered_exhibitors = [
                exhibitor for exhibitor in viz_data['df'].to_dict('records')
                if selected_category in exhibitor['categories']
            ]
        
        # Sort by company name
        filtered_exhibitors = sorted(filtered_exhibitors, key=lambda x: x['company_name'])
        
        # Store filtered exhibitors for download
        app.selected_category_exhibitors = {exhibitor['company_name']: exhibitor for exhibitor in filtered_exhibitors}
        
        # 언어에 따라 설명 필드 선택
        description_field = 'description_ko' if current_language == 'ko' else 'description'
        
        # 언어 토글 버튼 스타일 설정
        ko_style = {
            'marginRight': '10px', 
            'padding': '8px 15px', 
            'backgroundColor': '#4CAF50' if current_language == 'ko' else '#ccc', 
            'color': 'white' if current_language == 'ko' else 'black', 
            'border': 'none', 
            'borderRadius': '5px', 
            'cursor': 'pointer',
            'fontWeight': 'bold',
            'boxShadow': '0 2px 5px rgba(0,0,0,0.2)' if current_language == 'ko' else '0 2px 5px rgba(0,0,0,0.1)',
            'transition': 'all 0.3s ease'
        }
        
        en_style = {
            'padding': '8px 15px', 
            'backgroundColor': '#4CAF50' if current_language == 'en' else '#ccc', 
            'color': 'white' if current_language == 'en' else 'black', 
            'border': 'none', 
            'borderRadius': '5px', 
            'cursor': 'pointer',
            'fontWeight': 'bold',
            'boxShadow': '0 2px 5px rgba(0,0,0,0.2)' if current_language == 'en' else '0 2px 5px rgba(0,0,0,0.1)',
            'transition': 'all 0.3s ease'
        }
        
        # Create table with header and checkboxes
        return html.Div([
            html.H3(f'Companies in category: {selected_category} ({len(filtered_exhibitors)} exhibitors)', 
                   style={'textAlign': 'left', 'color': '#2C3E50', 'fontFamily': 'Arial, sans-serif', 'marginBottom': '15px', 'fontWeight': 'bold', 'borderBottom': '2px solid #4CAF50', 'paddingBottom': '10px'}),
            # 언어 토글 버튼 추가
            html.Div([
                html.Button('한글', id='ko-button-category', n_clicks=0, style=ko_style),
                html.Button('English', id='en-button-category', n_clicks=0, style=en_style)
            ], style={'marginBottom': '20px', 'textAlign': 'right', 'padding': '10px 0'}),
            html.Table(
                [html.Tr([
                    html.Th('Select', style={'width': '5%', 'padding': '12px', 'backgroundColor': '#4CAF50', 'color': 'white', 'textAlign': 'center', 'fontWeight': 'bold', 'borderBottom': '2px solid #ddd'}),
                    html.Th('Company Name', style={'width': '15%', 'padding': '12px', 'backgroundColor': '#4CAF50', 'color': 'white', 'textAlign': 'left', 'fontWeight': 'bold', 'borderBottom': '2px solid #ddd'}), 
                    html.Th('Booth Location', style={'width': '10%', 'padding': '12px', 'backgroundColor': '#4CAF50', 'color': 'white', 'textAlign': 'center', 'fontWeight': 'bold', 'borderBottom': '2px solid #ddd'}),
                    html.Th('Description', style={'width': '60%', 'padding': '12px', 'backgroundColor': '#4CAF50', 'color': 'white', 'textAlign': 'left', 'fontWeight': 'bold', 'borderBottom': '2px solid #ddd'}), 
                    html.Th('Homepage', style={'width': '10%', 'padding': '12px', 'backgroundColor': '#4CAF50', 'color': 'white', 'textAlign': 'center', 'fontWeight': 'bold', 'borderBottom': '2px solid #ddd'})
                ], style={'backgroundColor': '#f2f2f2'})] +
                [html.Tr([
                    html.Td(dcc.Checklist(
                        id={'type': 'category-checkbox', 'index': exhibitor['company_name']},
                        options=[{'label': '', 'value': exhibitor['company_name']}],
                        value=[],
                        style={'margin': '0', 'padding': '0'}
                    ), style={'width': '5%', 'padding': '10px', 'textAlign': 'center', 'borderBottom': '1px solid #ddd'}),
                    html.Td(exhibitor['company_name'], style={'width': '15%', 'padding': '10px', 'textAlign': 'left', 'borderBottom': '1px solid #ddd', 'fontWeight': 'bold'}),
                    html.Td(exhibitor['booth_location'], style={'width': '10%', 'padding': '10px', 'textAlign': 'center', 'borderBottom': '1px solid #ddd'}),
                    html.Td(exhibitor[description_field][:200] + '...' if exhibitor[description_field] and len(exhibitor[description_field]) > 200 else exhibitor[description_field], style={'width': '60%', 'padding': '10px', 'textAlign': 'left', 'borderBottom': '1px solid #ddd', 'lineHeight': '1.5'}),
                    html.Td(html.A('Website', href=next((contact['url'] for contact in exhibitor.get('contact_details', []) if contact.get('type', '').lower() == 'website'), '#'), target='_blank', style={'textDecoration': 'none', 'color': '#4CAF50', 'fontWeight': 'bold'}) if any(contact.get('type', '').lower() == 'website' for contact in exhibitor.get('contact_details', [])) else '', style={'width': '10%', 'padding': '10px', 'textAlign': 'center', 'borderBottom': '1px solid #ddd'})
                ], style={'backgroundColor': i % 2 == 0 and '#f9f9f9' or 'white'}) for i, exhibitor in enumerate(filtered_exhibitors)],
                style={'width': '100%', 'borderCollapse': 'collapse', 'boxShadow': '0 4px 8px 0 rgba(0,0,0,0.2)', 'borderRadius': '5px', 'overflow': 'hidden', 'marginTop': '20px', 'fontFamily': 'Arial, sans-serif'}
            )
        ]), {'marginTop': '20px', 'padding': '12px 20px', 'backgroundColor': '#4CAF50', 'color': 'white', 'border': 'none', 'borderRadius': '5px', 'cursor': 'pointer', 'display': 'block', 'fontWeight': 'bold', 'boxShadow': '0 2px 5px rgba(0,0,0,0.2)', 'transition': 'background-color 0.3s'}
    
    @callback(
        Output('pavilion-click-output', 'children'),
        Output('download-pavilion-btn', 'style'),
        Input('pavilion-chart', 'clickData'),
        Input('language-store', 'data'),
        prevent_initial_call=False
    )
    def display_pavilion_click_data(clickData, language_data):
        # 어떤 입력이 콜백을 트리거했는지 확인
        triggered_id = ctx.triggered_id
        
        if not clickData:
            return html.P('Click on a pavilion bar to see exhibitors in that pavilion'), {'display': 'none'}
        
        # 현재 선택된 언어 가져오기
        current_language = language_data.get('language', 'ko')
        
        # Get the clicked pavilion
        selected_pavilion = clickData['points'][0]['y']
        
        # Filter exhibitors by selected pavilion
        filtered_exhibitors = [
            exhibitor for exhibitor in viz_data['df'].to_dict('records')
            if exhibitor['pavilion'] == selected_pavilion
        ]
        
        # Sort by company name
        filtered_exhibitors = sorted(filtered_exhibitors, key=lambda x: x['company_name'])
        
        # Store filtered exhibitors for download
        app.selected_pavilion_exhibitors = {exhibitor['company_name']: exhibitor for exhibitor in filtered_exhibitors}
        
        # 언어에 따라 설명 필드 선택
        description_field = 'description_ko' if current_language == 'ko' else 'description'
        
        # 언어 토글 버튼 스타일 설정
        ko_style = {
            'marginRight': '10px', 
            'padding': '8px 15px', 
            'backgroundColor': '#4CAF50' if current_language == 'ko' else '#ccc', 
            'color': 'white' if current_language == 'ko' else 'black', 
            'border': 'none', 
            'borderRadius': '5px', 
            'cursor': 'pointer',
            'fontWeight': 'bold',
            'boxShadow': '0 2px 5px rgba(0,0,0,0.2)' if current_language == 'ko' else '0 2px 5px rgba(0,0,0,0.1)',
            'transition': 'all 0.3s ease'
        }
        
        en_style = {
            'padding': '8px 15px', 
            'backgroundColor': '#4CAF50' if current_language == 'en' else '#ccc', 
            'color': 'white' if current_language == 'en' else 'black', 
            'border': 'none', 
            'borderRadius': '5px', 
            'cursor': 'pointer',
            'fontWeight': 'bold',
            'boxShadow': '0 2px 5px rgba(0,0,0,0.2)' if current_language == 'en' else '0 2px 5px rgba(0,0,0,0.1)',
            'transition': 'all 0.3s ease'
        }
        
        # Create table with header and checkboxes
        return html.Div([
            html.H3(f'Companies in pavilion: {selected_pavilion} ({len(filtered_exhibitors)} exhibitors)', 
                   style={'textAlign': 'left', 'color': '#2C3E50', 'fontFamily': 'Arial, sans-serif', 'marginBottom': '15px', 'fontWeight': 'bold', 'borderBottom': '2px solid #4CAF50', 'paddingBottom': '10px'}),
            # 언어 토글 버튼 추가 (카테고리와 다른 ID 사용)
            html.Div([
                html.Button('한글', id='ko-button-pavilion', n_clicks=0, style=ko_style),
                html.Button('English', id='en-button-pavilion', n_clicks=0, style=en_style)
            ], style={'marginBottom': '20px', 'textAlign': 'right', 'padding': '10px 0'}),
            html.Table(
                [html.Tr([
                    html.Th('Select', style={'width': '5%', 'padding': '12px', 'backgroundColor': '#4CAF50', 'color': 'white', 'textAlign': 'center', 'fontWeight': 'bold', 'borderBottom': '2px solid #ddd'}),
                    html.Th('Company Name', style={'width': '15%', 'padding': '12px', 'backgroundColor': '#4CAF50', 'color': 'white', 'textAlign': 'left', 'fontWeight': 'bold', 'borderBottom': '2px solid #ddd'}), 
                    html.Th('Booth Location', style={'width': '10%', 'padding': '12px', 'backgroundColor': '#4CAF50', 'color': 'white', 'textAlign': 'center', 'fontWeight': 'bold', 'borderBottom': '2px solid #ddd'}),
                    html.Th('Description', style={'width': '60%', 'padding': '12px', 'backgroundColor': '#4CAF50', 'color': 'white', 'textAlign': 'left', 'fontWeight': 'bold', 'borderBottom': '2px solid #ddd'}), 
                    html.Th('Homepage', style={'width': '10%', 'padding': '12px', 'backgroundColor': '#4CAF50', 'color': 'white', 'textAlign': 'center', 'fontWeight': 'bold', 'borderBottom': '2px solid #ddd'})
                ], style={'backgroundColor': '#f2f2f2'})] +
                [html.Tr([
                    html.Td(dcc.Checklist(
                        id={'type': 'pavilion-checkbox', 'index': exhibitor['company_name']},
                        options=[{'label': '', 'value': exhibitor['company_name']}],
                        value=[],
                        style={'margin': '0', 'padding': '0'}
                    ), style={'width': '5%', 'padding': '10px', 'textAlign': 'center', 'borderBottom': '1px solid #ddd'}),
                    html.Td(exhibitor['company_name'], style={'width': '15%', 'padding': '10px', 'textAlign': 'left', 'borderBottom': '1px solid #ddd', 'fontWeight': 'bold'}),
                    html.Td(exhibitor['booth_location'], style={'width': '10%', 'padding': '10px', 'textAlign': 'center', 'borderBottom': '1px solid #ddd'}),
                    html.Td(exhibitor[description_field][:200] + '...' if exhibitor[description_field] and len(exhibitor[description_field]) > 200 else exhibitor[description_field], style={'width': '60%', 'padding': '10px', 'textAlign': 'left', 'borderBottom': '1px solid #ddd', 'lineHeight': '1.5'}),
                    html.Td(html.A('Website', href=next((contact['url'] for contact in exhibitor.get('contact_details', []) if contact.get('type', '').lower() == 'website'), '#'), target='_blank', style={'textDecoration': 'none', 'color': '#4CAF50', 'fontWeight': 'bold'}) if any(contact.get('type', '').lower() == 'website' for contact in exhibitor.get('contact_details', [])) else '', style={'width': '10%', 'padding': '10px', 'textAlign': 'center', 'borderBottom': '1px solid #ddd'})
                ], style={'backgroundColor': i % 2 == 0 and '#f9f9f9' or 'white'}) for i, exhibitor in enumerate(filtered_exhibitors)],
                style={'width': '100%', 'borderCollapse': 'collapse', 'boxShadow': '0 4px 8px 0 rgba(0,0,0,0.2)', 'borderRadius': '5px', 'overflow': 'hidden', 'marginTop': '20px', 'fontFamily': 'Arial, sans-serif'}
            )
        ]), {'marginTop': '20px', 'padding': '12px 20px', 'backgroundColor': '#4CAF50', 'color': 'white', 'border': 'none', 'borderRadius': '5px', 'cursor': 'pointer', 'display': 'block', 'fontWeight': 'bold', 'boxShadow': '0 2px 5px rgba(0,0,0,0.2)', 'transition': 'background-color 0.3s'}
    
    @callback(
        Output('download-category-data', 'data'),
        Input('download-category-btn', 'n_clicks'),
        State({'type': 'category-checkbox', 'index': ALL}, 'value'),
        State({'type': 'category-checkbox', 'index': ALL}, 'id'),
        prevent_initial_call=True
    )
    def download_selected_category_data(n_clicks, checkbox_values, checkbox_ids):
        if not n_clicks:
            return None
        
        # Get selected company names
        selected_companies = []
        for i, values in enumerate(checkbox_values):
            if values:  # If checklist has any values selected
                selected_companies.append(values[0])  # Get the first (and only) value
        
        if not selected_companies:
            return None
        
        # Create DataFrame with selected companies
        selected_data = []
        for company_name in selected_companies:
            if company_name in app.selected_category_exhibitors:
                selected_data.append(app.selected_category_exhibitors[company_name])
        
        if not selected_data:
            return None
        
        # Convert to DataFrame
        df = pd.DataFrame(selected_data)
        
        # 필요한 필드만 선택
        if not df.empty:
            df = df[['company_name', 'booth_location', 'pavilion', 'description', 'description_ko', 'website']]
            # 컬럼명 변경
            df.columns = ['회사명', '부스 위치', '파빌리온', '영문 설명', '한글 설명', '웹페이지']
        
        # Return Excel file
        return dcc.send_data_frame(df.to_excel, "selected_companies.xlsx", sheet_name="Selected Companies")
    
    @callback(
        Output('download-pavilion-data', 'data'),
        Input('download-pavilion-btn', 'n_clicks'),
        State({'type': 'pavilion-checkbox', 'index': ALL}, 'value'),
        State({'type': 'pavilion-checkbox', 'index': ALL}, 'id'),
        prevent_initial_call=True
    )
    def download_selected_pavilion_data(n_clicks, checkbox_values, checkbox_ids):
        if not n_clicks:
            return None
        
        # Get selected company names
        selected_companies = []
        for i, values in enumerate(checkbox_values):
            if values:  # If checklist has any values selected
                selected_companies.append(values[0])  # Get the first (and only) value
        
        if not selected_companies:
            return None
        
        # Create DataFrame with selected companies
        selected_data = []
        for company_name in selected_companies:
            if company_name in app.selected_pavilion_exhibitors:
                selected_data.append(app.selected_pavilion_exhibitors[company_name])
        
        if not selected_data:
            return None
        
        # Convert to DataFrame
        df = pd.DataFrame(selected_data)
        
        # 필요한 필드만 선택
        if not df.empty:
            df = df[['company_name', 'booth_location', 'pavilion', 'description', 'description_ko', 'website']]
            # 컬럼명 변경
            df.columns = ['회사명', '부스 위치', '파빌리온', '영문 설명', '한글 설명', '웹페이지']
        
        # Return Excel file
        return dcc.send_data_frame(df.to_excel, "selected_companies.xlsx", sheet_name="Selected Companies")
    
    @callback(
        Output('language-store', 'data', allow_duplicate=True),
        Output('ko-button-pavilion', 'style'),
        Output('en-button-pavilion', 'style'),
        Input('ko-button-pavilion', 'n_clicks'),
        Input('en-button-pavilion', 'n_clicks'),
        State('language-store', 'data'),
        prevent_initial_call=True
    )
    def toggle_language_pavilion(ko_clicks, en_clicks, language_data):
        # 어떤 버튼이 클릭되었는지 확인
        triggered_id = ctx.triggered_id
        
        # 한글 버튼 스타일
        ko_style = {
            'marginRight': '10px', 
            'padding': '8px 15px', 
            'backgroundColor': '#4CAF50', 
            'color': 'white', 
            'border': 'none', 
            'borderRadius': '5px', 
            'cursor': 'pointer',
            'fontWeight': 'bold',
            'boxShadow': '0 2px 5px rgba(0,0,0,0.2)',
            'transition': 'all 0.3s ease'
        }
        
        # 영어 버튼 스타일
        en_style = {
            'padding': '8px 15px', 
            'backgroundColor': '#ccc', 
            'color': 'black', 
            'border': 'none', 
            'borderRadius': '5px', 
            'cursor': 'pointer',
            'fontWeight': 'bold',
            'boxShadow': '0 2px 5px rgba(0,0,0,0.1)',
            'transition': 'all 0.3s ease'
        }
        
        # 비활성화된 한글 버튼 스타일
        ko_inactive_style = {
            'marginRight': '10px', 
            'padding': '8px 15px', 
            'backgroundColor': '#ccc', 
            'color': 'black', 
            'border': 'none', 
            'borderRadius': '5px', 
            'cursor': 'pointer',
            'fontWeight': 'bold',
            'boxShadow': '0 2px 5px rgba(0,0,0,0.1)',
            'transition': 'all 0.3s ease'
        }
        
        # 비활성화된 영어 버튼 스타일
        en_inactive_style = {
            'padding': '8px 15px', 
            'backgroundColor': '#4CAF50', 
            'color': 'white', 
            'border': 'none', 
            'borderRadius': '5px', 
            'cursor': 'pointer',
            'fontWeight': 'bold',
            'boxShadow': '0 2px 5px rgba(0,0,0,0.2)',
            'transition': 'all 0.3s ease'
        }
        
        if triggered_id == 'ko-button-pavilion':
            return {'language': 'ko'}, ko_style, ko_inactive_style
        else:  # en-button-pavilion
            return {'language': 'en'}, en_inactive_style, en_style
    
    return app

def main():
    # Load and process data
    data = load_data()
    processed_data = process_data(data)
    viz_data = create_visualizations(processed_data)
    
    # Create and run the app
    app = create_app(viz_data)
    
    # Get port from environment variable or use default
    port = int(os.environ.get('PORT', 8050))
    
    if __name__ == '__main__':
        print(f"Starting Dash server on port {port}")
        app.run_server(host='0.0.0.0', port=port, debug=False)
    
    return app

# Create the app instance
app = main()
server = app.server  # Expose server variable for render.com

if __name__ == "__main__":
    app.run_server(debug=False)
