from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.chart import XL_CHART_TYPE
from pptx.enum.dml import MSO_THEME_COLOR
from pptx.chart.data import CategoryChartData
import os
from datetime import datetime
import comtypes.client
import comtypes

def _get_theme_colors(theme_name):
    # Define theme-specific color palettes
    color_palettes = {
        "Professional": {
            "primary": RGBColor(0, 112, 192),    # Professional blue
            "secondary": RGBColor(0, 176, 240),  # Light blue
            "accent": RGBColor(255, 255, 255),   # White
            "text": RGBColor(68, 68, 68)         # Dark gray
        },
        "Creative": {
            "primary": RGBColor(255, 89, 94),    # Vibrant red
            "secondary": RGBColor(255, 202, 58), # Bright yellow
            "accent": RGBColor(138, 201, 38),    # Fresh green
            "text": RGBColor(25, 25, 25)         # Near black
        },
        "Corporate": {
            "primary": RGBColor(31, 73, 125),    # Deep blue
            "secondary": RGBColor(79, 129, 189), # Medium blue
            "accent": RGBColor(192, 80, 77),     # Corporate red
            "text": RGBColor(89, 89, 89)         # Medium gray
        },
        "Modern": {
            "primary": RGBColor(45, 55, 72),     # Dark slate
            "secondary": RGBColor(237, 242, 247),# Light gray
            "accent": RGBColor(66, 153, 225),    # Modern blue
            "text": RGBColor(26, 32, 44)         # Dark slate
        },
        "Elegant": {
            "primary": RGBColor(72, 52, 52),     # Deep burgundy
            "secondary": RGBColor(192, 152, 152),# Soft rose
            "accent": RGBColor(240, 230, 220),   # Cream
            "text": RGBColor(51, 51, 51)         # Rich black
        }
    }
    return color_palettes.get(theme_name, color_palettes["Professional"])

def _apply_theme(prs, theme_name):
    # Enhanced theme configurations with specific requirements
    theme_map = {
        "Professional": {
            "layout": "Office Theme",
            "elements": ["transition", "minimal_shape"],
            "title_font_size": 40,
            "subtitle_font_size": 28,
            "content_font_size": 20,
            "bullet_font_size": 18,
            "shape_types": [
                MSO_SHAPE.ROUNDED_RECTANGLE,
                MSO_SHAPE.RECTANGLE,
                MSO_SHAPE.FLOWCHART_PROCESS,
                MSO_SHAPE.FLOWCHART_DECISION,
                MSO_SHAPE.BLOCK_ARC
            ],
            "chart_types": [XL_CHART_TYPE.COLUMN_CLUSTERED],
            "design_elements": ["clean_lines", "subtle_shadows"],
            "transition": {
                "effect": "fade",
                "duration": 0.5
            }
        },
        "Creative": {
            "layout": "Facet",
            "elements": ["transition", "shape", "artistic"],
            "title_font_size": 44,
            "subtitle_font_size": 32,
            "content_font_size": 24,
            "bullet_font_size": 22,
            "shape_types": [
                MSO_SHAPE.STAR_5_POINT,
                MSO_SHAPE.STAR_8_POINT,
                MSO_SHAPE.CLOUD,
                MSO_SHAPE.SUN,
                MSO_SHAPE.WAVE,
                MSO_SHAPE.DOUBLE_WAVE,
                MSO_SHAPE.BALLOON
            ],
            "chart_types": [XL_CHART_TYPE.PIE, XL_CHART_TYPE.DOUGHNUT],
            "design_elements": ["gradients", "overlapping_shapes", "dynamic_lines"],
            "transition": {
                "effect": "wipe",
                "duration": 0.3
            }
        },
        "Corporate": {
            "layout": "Office Theme",
            "elements": ["chart", "shape", "transition", "data_visualization"],
            "title_font_size": 36,
            "subtitle_font_size": 28,
            "content_font_size": 20,
            "bullet_font_size": 18,
            "shape_types": [
                MSO_SHAPE.FLOWCHART_PROCESS,
                MSO_SHAPE.FLOWCHART_DECISION,
                MSO_SHAPE.FLOWCHART_TERMINATOR,
                MSO_SHAPE.CUBE,
                MSO_SHAPE.BLOCK_ARC,
                MSO_SHAPE.ROUNDED_RECTANGLE
            ],
            "chart_types": [XL_CHART_TYPE.COLUMN_CLUSTERED, XL_CHART_TYPE.LINE, XL_CHART_TYPE.BAR_CLUSTERED],
            "design_elements": ["grid_lines", "professional_icons", "data_highlights"],
            "transition": {
                "effect": "push",
                "duration": 0.5
            }
        },
        "Modern": {
            "layout": "Ion",
            "elements": ["design", "shape", "transition", "minimal"],
            "title_font_size": 42,
            "subtitle_font_size": 30,
            "content_font_size": 22,
            "bullet_font_size": 20,
            "shape_types": [
                MSO_SHAPE.ROUNDED_RECTANGLE,
                MSO_SHAPE.CIRCULAR_ARROW,
                MSO_SHAPE.WAVE,
                MSO_SHAPE.DOUBLE_WAVE,
                MSO_SHAPE.BEVEL,
                MSO_SHAPE.FOLDED_CORNER
            ],
            "chart_types": [XL_CHART_TYPE.LINE_MARKERS, XL_CHART_TYPE.AREA],
            "design_elements": ["geometric_patterns", "bold_colors", "minimal_icons"],
            "transition": {
                "effect": "cut",
                "duration": 0.3
            }
        },
        "Elegant": {
            "layout": "Office Theme",
            "elements": ["design", "chart", "shape", "transition", "premium"],
            "title_font_size": 38,
            "subtitle_font_size": 28,
            "content_font_size": 20,
            "bullet_font_size": 18,
            "shape_types": [
                MSO_SHAPE.STAR_8_POINT,
                MSO_SHAPE.STAR_12_POINT,
                MSO_SHAPE.CIRCULAR_ARROW,
                MSO_SHAPE.BLOCK_ARC,
                MSO_SHAPE.CHORD
            ],
            "chart_types": [XL_CHART_TYPE.LINE, XL_CHART_TYPE.AREA_STACKED],
            "design_elements": ["subtle_patterns", "gold_accents", "sophisticated_icons"],
            "transition": {
                "effect": "dissolve",
                "duration": 0.7
            }
        }
    }
    return theme_map.get(theme_name, theme_map["Professional"])

def _apply_color_scheme(prs, color_scheme, theme_name):
    # Get theme-specific colors
    theme_colors = _get_theme_colors(theme_name)
    
    # Override with color scheme if specified
    color_schemes = {
        "Default": theme_colors,
        "Blue": {
            "primary": RGBColor(0, 112, 192),
            "secondary": RGBColor(0, 176, 240),
            "accent": RGBColor(255, 255, 255),
            "text": RGBColor(68, 68, 68)
        },
        "Green": {
            "primary": RGBColor(0, 176, 80),
            "secondary": RGBColor(146, 208, 80),
            "accent": RGBColor(255, 255, 255),
            "text": RGBColor(68, 68, 68)
        },
        "Red": {
            "primary": RGBColor(192, 0, 0),
            "secondary": RGBColor(255, 0, 0),
            "accent": RGBColor(255, 255, 255),
            "text": RGBColor(68, 68, 68)
        },
        "Purple": {
            "primary": RGBColor(112, 48, 160),
            "secondary": RGBColor(149, 55, 53),
            "accent": RGBColor(255, 255, 255),
            "text": RGBColor(68, 68, 68)
        },
        "Orange": {
            "primary": RGBColor(255, 140, 0),
            "secondary": RGBColor(255, 192, 0),
            "accent": RGBColor(255, 255, 255),
            "text": RGBColor(68, 68, 68)
        }
    }
    return color_schemes.get(color_scheme, theme_colors)

def _apply_font(shape, font_name, font_size=None, color=None, bold=False, italic=False):
    if hasattr(shape, "text_frame"):
        for paragraph in shape.text_frame.paragraphs:
            for run in paragraph.runs:
                run.font.name = font_name
                if font_size:
                    run.font.size = Pt(font_size)
                if color:
                    run.font.color.rgb = color
                if bold:
                    run.font.bold = True
                if italic:
                    run.font.italic = True

def _add_shape(slide, shape_type, left, top, width, height, fill_color=None, line_color=None):
    shape = slide.shapes.add_shape(
        shape_type,
        Inches(left),
        Inches(top),
        Inches(width),
        Inches(height)
    )
    if fill_color:
        shape.fill.solid()
        shape.fill.fore_color.rgb = fill_color
    if line_color:
        shape.line.color.rgb = line_color
    return shape

def _add_chart(slide, chart_type, left, top, width, height, data, theme_colors):
    # Create chart data object
    chart_data = CategoryChartData()
    
    # Add categories
    chart_data.categories = data.get("categories", ["Q1", "Q2", "Q3", "Q4"])
    
    # Add series
    for series_name, series_values in data.get("series", []):
        chart_data.add_series(series_name, series_values)
    
    # Add chart to slide
    chart = slide.shapes.add_chart(
        chart_type,
        Inches(left),
        Inches(top),
        Inches(width),
        Inches(height),
        chart_data
    ).chart
    
    # Apply theme colors to chart
    chart.chart_style = 2  # Use a clean style
    
    # Format chart elements
    if hasattr(chart, 'has_legend') and chart.has_legend:
        chart.legend.font.size = Pt(10)
        chart.legend.font.color.rgb = theme_colors["text"]
    
    if hasattr(chart, 'chart_title') and chart.chart_title:
        chart.chart_title.font.size = Pt(14)
        chart.chart_title.font.color.rgb = theme_colors["primary"]
    
    # Format series
    for idx, series in enumerate(chart.series):
        series.format.fill.solid()
        # Alternate between primary and secondary colors
        color = theme_colors["primary"] if idx % 2 == 0 else theme_colors["secondary"]
        series.format.fill.fore_color.rgb = color
    
    return chart

def _get_topic_based_elements(topic, theme_info, theme_colors):
    """Determine appropriate shapes and elements based on the presentation topic."""
    topic = topic.lower()
    
    # Define topic categories and their associated elements using only basic shapes
    topic_elements = {
        "business": {
            "shapes": [
                (MSO_SHAPE.RECTANGLE, "Business growth chart"),
                (MSO_SHAPE.FLOWCHART_PROCESS, "Business process"),
                (MSO_SHAPE.CUBE, "Business building blocks"),
                (MSO_SHAPE.ROUNDED_RECTANGLE, "Business card style")
            ],
            "positions": [
                (7, 1, 2, 1),  # right side
                (0.5, 0.5, 1.5, 0.5),  # top left
                (8, 2, 1, 1),  # bottom right
                (0, 5, 2, 0.3)  # bottom left
            ]
        },
        "technology": {
            "shapes": [
                (MSO_SHAPE.HEXAGON, "Tech network"),
                (MSO_SHAPE.ROUNDED_RECTANGLE, "Technology cycle"),
                (MSO_SHAPE.ROUNDED_RECTANGLE, "Tech interface"),
                (MSO_SHAPE.FLOWCHART_DECISION, "Tech decision point")
            ],
            "positions": [
                (7.5, 1, 1.5, 1.5),
                (0.5, 0.5, 1.2, 1.2),
                (8, 2, 1.5, 0.8),
                (0.5, 5, 1.5, 0.8)
            ]
        },
        "education": {
            "shapes": [
                (MSO_SHAPE.OVAL, "Learning cycle"),
                (MSO_SHAPE.ROUNDED_RECTANGLE, "Achievement"),
                (MSO_SHAPE.ROUNDED_RECTANGLE, "Book style"),
                (MSO_SHAPE.FLOWCHART_PROCESS, "Learning process")
            ],
            "positions": [
                (7, 1, 1.5, 1),
                (0.5, 0.5, 1, 1),
                (8, 2, 1.2, 1.5),
                (0.5, 5, 1.5, 0.5)
            ]
        },
        "health": {
            "shapes": [
                (MSO_SHAPE.OVAL, "Health symbol"),
                (MSO_SHAPE.OVAL, "Medical cycle"),
                (MSO_SHAPE.ROUNDED_RECTANGLE, "Medical card"),
                (MSO_SHAPE.FLOWCHART_PROCESS, "Health process")
            ],
            "positions": [
                (7.5, 1, 1, 1),
                (0.5, 0.5, 1.2, 1.2),
                (8, 2, 1.2, 0.8),
                (0.5, 5, 1.5, 0.5)
            ]
        },
        "environment": {
            "shapes": [
                (MSO_SHAPE.OVAL, "Sun energy"),
                (MSO_SHAPE.ROUNDED_RECTANGLE, "Cloud/atmosphere"),
                (MSO_SHAPE.ROUNDED_RECTANGLE, "Water/waves"),
                (MSO_SHAPE.OVAL, "Earth cycle")
            ],
            "positions": [
                (7.5, 0.5, 1, 1),
                (0.5, 0.5, 1.2, 0.8),
                (8, 2, 1.5, 0.5),
                (0.5, 5, 1.2, 1.2)
            ]
        },
        "finance": {
            "shapes": [
                (MSO_SHAPE.RECTANGLE, "Financial chart"),
                (MSO_SHAPE.ROUNDED_RECTANGLE, "Financial cycle"),
                (MSO_SHAPE.CUBE, "Investment blocks"),
                (MSO_SHAPE.ROUNDED_RECTANGLE, "Financial card")
            ],
            "positions": [
                (7, 1, 2, 1),
                (0.5, 0.5, 1.2, 1.2),
                (8, 2, 1, 1),
                (0.5, 5, 1.5, 0.3)
            ]
        }
    }
    
    # Default elements if no specific topic is matched
    default_elements = {
        "shapes": [
            (MSO_SHAPE.ROUNDED_RECTANGLE, "Decorative element"),
            (MSO_SHAPE.OVAL, "Design element"),
            (MSO_SHAPE.FLOWCHART_PROCESS, "Process element"),
            (MSO_SHAPE.ROUNDED_RECTANGLE, "Cycle element")
        ],
        "positions": [
            (7, 1, 2, 0.5),
            (0.5, 0.5, 1.2, 1.2),
            (8, 2, 1.5, 0.5),
            (0.5, 5, 1.5, 0.5)
        ]
    }
    
    # Determine the most relevant topic category
    matched_topic = None
    for category in topic_elements:
        if category in topic:
            matched_topic = category
            break
    
    # Get elements based on matched topic or use default
    elements = topic_elements.get(matched_topic, default_elements)
    
    return elements["shapes"], elements["positions"]

def _apply_theme_elements(slide, theme_info, formatting, theme_colors, topic=""):
    if not formatting or "elements" not in formatting:
        return

    for element in formatting.get("elements", []):
        element_type = element.get("type", "")
        
        if element_type == "shape" and "shape" in theme_info["elements"]:
            # Get topic-based elements
            shapes, positions = _get_topic_based_elements(topic, theme_info, theme_colors)
            
            # Apply shapes based on theme and topic
            for (shape_type, _), (left, top, width, height) in zip(shapes, positions):
                # Adjust colors based on theme
                if theme_info["layout"] == "Facet":  # Creative theme
                    fill_color = theme_colors["accent"] if shape_type in [MSO_SHAPE.SUN, MSO_SHAPE.CLOUD] else theme_colors["primary"]
                    line_color = theme_colors["primary"] if shape_type in [MSO_SHAPE.STAR_8_POINT, MSO_SHAPE.HEART] else None
                elif theme_info["layout"] == "Ion":  # Modern theme
                    fill_color = theme_colors["secondary"] if shape_type in [MSO_SHAPE.CIRCULAR_ARROW, MSO_SHAPE.WAVE] else theme_colors["primary"]
                    line_color = None
                elif "Elegant" in theme_info["layout"]:  # Elegant theme
                    fill_color = theme_colors["accent"] if shape_type in [MSO_SHAPE.STAR_12_POINT, MSO_SHAPE.OVAL] else theme_colors["primary"]
                    line_color = theme_colors["primary"] if shape_type in [MSO_SHAPE.STAR_12_POINT, MSO_SHAPE.CIRCULAR_ARROW] else None
                else:  # Professional and Corporate themes
                    fill_color = theme_colors["primary"]
                    line_color = None
                
                _add_shape(slide, shape_type, left, top, width, height,
                          fill_color=fill_color,
                          line_color=line_color)
        
        elif element_type == "chart" and "chart" in theme_info["elements"]:
            # Create sample chart data based on topic
            chart_data = _get_topic_based_chart_data(topic)
            chart_type = theme_info["chart_types"][0]
            _add_chart(slide, chart_type, 5, 2, 4, 3, chart_data, theme_colors)
        
        elif element_type == "design" and "design" in theme_info["elements"]:
            # Add design elements based on theme and topic
            if theme_info["layout"] == "Ion":  # Modern theme
                _add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, 0, 0, 10, 0.5,
                          fill_color=theme_colors["primary"])
                _add_shape(slide, MSO_SHAPE.OVAL, 7, 1, 2, 0.3,
                          fill_color=theme_colors["secondary"])
            elif "Elegant" in theme_info["layout"]:  # Elegant theme
                _add_shape(slide, MSO_SHAPE.OVAL, 7, 1, 1, 1,
                          fill_color=theme_colors["accent"])
                _add_shape(slide, MSO_SHAPE.OVAL, 8, 0, 1, 1,
                          fill_color=theme_colors["secondary"])

def _get_topic_based_chart_data(topic):
    """Generate appropriate chart data based on the presentation topic."""
    topic = topic.lower()
    
    # Define topic-specific chart data
    chart_data = {
        "business": {
            "categories": ["Q1", "Q2", "Q3", "Q4"],
            "series": [
                ("Revenue", [30, 45, 35, 50]),
                ("Growth", [40, 55, 45, 60])
            ]
        },
        "technology": {
            "categories": ["2020", "2021", "2022", "2023"],
            "series": [
                ("Adoption", [20, 35, 50, 75]),
                ("Innovation", [30, 45, 60, 85])
            ]
        },
        "education": {
            "categories": ["Year 1", "Year 2", "Year 3", "Year 4"],
            "series": [
                ("Performance", [60, 75, 85, 90]),
                ("Progress", [50, 65, 80, 95])
            ]
        },
        "health": {
            "categories": ["Jan", "Apr", "Jul", "Oct"],
            "series": [
                ("Wellness", [70, 75, 80, 85]),
                ("Recovery", [60, 70, 75, 80])
            ]
        },
        "environment": {
            "categories": ["2019", "2020", "2021", "2022"],
            "series": [
                ("Conservation", [40, 50, 60, 70]),
                ("Impact", [30, 45, 55, 65])
            ]
        },
        "finance": {
            "categories": ["Q1", "Q2", "Q3", "Q4"],
            "series": [
                ("Investment", [25, 35, 45, 55]),
                ("Returns", [20, 30, 40, 50])
            ]
        }
    }
    
    # Default chart data if no specific topic is matched
    default_data = {
        "categories": ["Q1", "Q2", "Q3", "Q4"],
        "series": [
            ("Series 1", [30, 45, 35, 50]),
            ("Series 2", [40, 55, 45, 60])
        ]
    }
    
    # Find matching topic data or use default
    for category in chart_data:
        if category in topic:
            return chart_data[category]
    
    return default_data

def _apply_comtypes_transitions(filepath, theme_name):
    """Apply slide transitions using comtypes based on the theme."""
    # Define theme-specific transitions
    theme_transitions = {
        "Professional": {"effect": "fade", "duration": 0.5},
        "Creative": {"effect": "wipe", "duration": 0.3},
        "Corporate": {"effect": "push", "duration": 0.5},
        "Modern": {"effect": "cut", "duration": 0.3},
        "Elegant": {"effect": "dissolve", "duration": 0.7}
    }

    # Get the transition settings for the theme, default to Professional
    transition = theme_transitions.get(theme_name, theme_transitions["Professional"])
    effect = transition["effect"]
    duration = transition["duration"]

    comtypes.CoInitialize()  # Initialize COM
    try:
        powerpoint = comtypes.client.CreateObject("PowerPoint.Application")
        powerpoint.Visible = 1

        presentation = powerpoint.Presentations.Open(filepath)

        for slide in presentation.Slides:
            slide.SlideShowTransition.EntryEffect = getattr(
                comtypes.gen.PowerPoint.PpEntryEffect,
                f"ppEffect{effect.capitalize()}",
                0
            )
            slide.SlideShowTransition.Duration = duration

        presentation.Save()
        presentation.Close()
        powerpoint.Quit()
    finally:
        comtypes.CoUninitialize()  # Uninitialize COM

def generate_ppt_doc(data):
    prs = Presentation()
    
    # Get the presentation topic from the title
    topic = data.get("title", "").lower()
    
    # Apply theme and color scheme
    theme_info = _apply_theme(prs, data.get("theme", "Professional"))
    theme_colors = _apply_color_scheme(prs, data.get("color_scheme", "Default"), data.get("theme", "Professional"))
    font_name = data.get("font", "Calibri")
    
    # Create title slide
    title_slide_layout = prs.slide_layouts[0]
    slide = prs.slides.add_slide(title_slide_layout)
    
    # Apply title and subtitle with theme-specific formatting
    title_shape = slide.shapes.title
    title_shape.text = data.get("title", "Untitled Presentation")
    _apply_font(title_shape, font_name, 
                theme_info["title_font_size"],
                theme_colors["primary"],
                bold=True)
    
    subtitle_shape = slide.placeholders[1]
    subtitle_shape.text = data.get("subtitle", "")
    _apply_font(subtitle_shape, font_name,
                theme_info["subtitle_font_size"],
                theme_colors["secondary"])
    
    # Process slides
    slides_data = data.get("slides", [])
    include_toc = data.get("include_toc", True)
    include_notes = data.get("include_notes", True)
    
    # Add slides
    for idx, slide_data in enumerate(slides_data):
        # Choose layout based on whether it's a TOC slide
        is_toc = slide_data.get("title", "").lower() == "table of contents"
        
        if is_toc:
            # Use a custom layout for TOC
            layout = prs.slide_layouts[2]  # Using a blank layout for TOC
            slide = prs.slides.add_slide(layout)
            
            # Add title with proper spacing
            title_shape = slide.shapes.title
            title_shape.text = "Table of Contents"
            _apply_font(title_shape, font_name,
                        theme_info["title_font_size"],
                        theme_colors["primary"],
                        bold=True)
            
            # Create a text box for TOC content with improved positioning
            left = Inches(1.5)  # Increased left margin
            top = Inches(2.2)   # Adjusted top position
            width = Inches(7)   # Adjusted width for better alignment
            height = Inches(4)
            
            toc_box = slide.shapes.add_textbox(left, top, width, height)
            tf = toc_box.text_frame
            tf.word_wrap = True
            tf.margin_left = Inches(0.1)  # Add small left margin
            tf.margin_right = Inches(0.1) # Add small right margin
            
            # Process TOC content
            content = slide_data.get("content", "")
            lines = content.split('\n')
            
            # Add each line as a paragraph with proper formatting
            for line in lines:
                if line.strip():
                    # Clean the line of any asterisks or unwanted characters
                    clean_line = line.strip().replace('*', '').strip()
                    if not clean_line:
                        continue
                        
                    p = tf.add_paragraph()
                    p.text = clean_line
                    p.level = 0  # Main level
                    p.alignment = PP_ALIGN.LEFT
                    p.space_before = Pt(6)  # Add space before paragraph
                    p.space_after = Pt(6)   # Add space after paragraph
                    
                    # Format the paragraph
                    for run in p.runs:
                        run.font.name = font_name
                        run.font.size = Pt(theme_info["content_font_size"])
                        run.font.color.rgb = theme_colors["text"]
                        
                        # Make numbers and titles bold
                        if clean_line[0].isdigit() or clean_line.startswith(('Chapter', 'Section', 'Part')):
                            run.font.bold = True
                            run.font.color.rgb = theme_colors["primary"]
            
            # Add decorative elements for TOC with improved positioning
            if theme_info["layout"] == "Facet":  # Creative theme
                _add_shape(slide, MSO_SHAPE.OVAL, 0.8, 1.8, 0.3, 0.3,
                          fill_color=theme_colors["accent"])
                _add_shape(slide, MSO_SHAPE.OVAL, 8.9, 1.8, 0.3, 0.3,
                          fill_color=theme_colors["accent"])
            elif theme_info["layout"] == "Ion":  # Modern theme
                _add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, 0.5, 2, 9, 0.1,
                          fill_color=theme_colors["primary"])
                _add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, 0.5, 5.8, 9, 0.1,
                          fill_color=theme_colors["primary"])
            else:  # Professional and other themes
                _add_shape(slide, MSO_SHAPE.RECTANGLE, 0.5, 2, 9, 0.05,
                          fill_color=theme_colors["primary"])
                _add_shape(slide, MSO_SHAPE.RECTANGLE, 0.5, 5.8, 9, 0.05,
                          fill_color=theme_colors["primary"])
            
        else:
            # Regular slide
            layout = prs.slide_layouts[1]
            slide = prs.slides.add_slide(layout)
            
            # Get slide-specific formatting
            formatting = slide_data.get("formatting", {})
            title_font_size = formatting.get("title_font_size", theme_info["title_font_size"])
            content_font_size = formatting.get("content_font_size", theme_info["content_font_size"])
            
            # Apply title
            title_shape = slide.shapes.title
            title_shape.text = slide_data.get("title", "")
            _apply_font(title_shape, font_name,
                        title_font_size,
                        theme_colors["primary"],
                        bold=True)
            
            # Apply content with proper formatting
            content_shape = slide.placeholders[1]
            content_shape.text = slide_data.get("content", "")
            _apply_font(content_shape, font_name,
                        content_font_size,
                        theme_colors["text"])
        
        # Apply theme-specific elements
        _apply_theme_elements(slide, theme_info, formatting, theme_colors, topic)
        
        # Apply speaker notes if enabled
        if include_notes and slide_data.get("notes"):
            notes_slide = slide.notes_slide
            notes_slide.notes_text_frame.text = slide_data.get("notes", "")
            _apply_font(notes_slide.notes_text_frame, font_name, 12,
                        theme_colors["text"])
        
        # Apply slide transition
        if hasattr(slide, 'transition'):
            # Use different transitions based on theme and slide type
            if is_toc:
                # Special transition for TOC
                slide.transition.transition_type = 'fade'
                slide.transition.transition_speed = 0.5
            else:
                # Theme-specific transitions
                if theme_info["layout"] == "Facet":  # Creative theme
                    transitions = ['wipe', 'push', 'cut', 'fade']
                elif theme_info["layout"] == "Ion":  # Modern theme
                    transitions = ['cut', 'wipe', 'push', 'fade']
                elif "Elegant" in theme_info["layout"]:  # Elegant theme
                    transitions = ['dissolve', 'fade', 'wipe', 'push']
                else:  # Professional and Corporate themes
                    transitions = ['fade', 'push', 'wipe', 'cut']
                
                # Cycle through transitions for variety
                transition = transitions[idx % len(transitions)]
                slide.transition.transition_type = transition
                slide.transition.transition_speed = 0.5

    # Generate a clean, descriptive filename
    title = data.get("title", "Untitled").strip()
    # Remove special characters and replace spaces with underscores
    clean_title = "".join(c for c in title if c.isalnum() or c.isspace())
    clean_title = clean_title.replace(" ", "_")
    # Limit title length and add timestamp
    if len(clean_title) > 30:
        clean_title = clean_title[:30]
    timestamp = datetime.now().strftime("%Y%m%d_%H%M")
    filename = f"{clean_title}_{timestamp}.pptx"
    
    # Save the presentation to Downloads folder
    downloads_path = os.path.expanduser("~/Downloads")
    filepath = os.path.join(downloads_path, filename)
    
    # Ensure Downloads folder exists
    os.makedirs(downloads_path, exist_ok=True)
    
    # Save the presentation
    prs.save(filepath)

    # Apply transitions using comtypes
    transition_effect = data.get("transition_effect", "fade")
    transition_duration = data.get("transition_duration", 1)
    _apply_comtypes_transitions(filepath, transition_effect, transition_duration)

    return filepath

