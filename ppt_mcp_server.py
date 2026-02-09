#!/usr/bin/env python
"""
MCP Server for PowerPoint manipulation using pywin32 (COM automation).
Enables LIVE editing of open PowerPoint presentations with visual feedback via screenshots.
"""
import os
import sys
import tempfile
import base64
from pathlib import Path
from typing import Dict, List, Optional, Any
from mcp.server.fastmcp import FastMCP

# Windows COM automation
import win32com.client
import pythoncom
from PIL import Image
import io

# Initialize the FastMCP server
app = FastMCP(
    name="ppt-live-server",
    instructions="MCP Server for LIVE PowerPoint manipulation using COM automation with screenshot capability",
)

# Global state
_ppt_app = None
_initialized = False


def get_ppt_app():
    """Get or create PowerPoint application instance."""
    global _ppt_app, _initialized

    pythoncom.CoInitialize()

    if _ppt_app is None:
        try:
            # Try to connect to existing PowerPoint instance
            _ppt_app = win32com.client.GetActiveObject("PowerPoint.Application")
        except:
            # Create new PowerPoint instance
            _ppt_app = win32com.client.Dispatch("PowerPoint.Application")

        _ppt_app.Visible = True
        _initialized = True

    return _ppt_app


def get_active_presentation():
    """Get the active presentation or raise error."""
    ppt = get_ppt_app()

    if ppt.Presentations.Count == 0:
        raise ValueError("No presentation is open. Use create_presentation or open_presentation first.")

    return ppt.ActivePresentation


# ============================================================================
# SCREENSHOT TOOLS - Most important for visual debugging
# ============================================================================

@app.tool()
def screenshot_slide(
    slide_number: int = 1,
    width: int = 1280,
    height: int = 720
) -> Dict:
    """
    Take a screenshot of a specific slide. Returns base64 encoded PNG image.

    This is the PRIMARY debugging tool - use it to visually verify any changes made.

    Args:
        slide_number: Slide number (1-based)
        width: Export width in pixels
        height: Export height in pixels

    Returns:
        Dict with base64 image data and metadata
    """
    try:
        pres = get_active_presentation()

        if slide_number < 1 or slide_number > pres.Slides.Count:
            return {
                "error": f"Invalid slide number: {slide_number}. Presentation has {pres.Slides.Count} slides."
            }

        slide = pres.Slides(slide_number)

        # Export slide to temporary file
        temp_file = os.path.join(tempfile.gettempdir(), f"ppt_screenshot_{slide_number}.png")
        slide.Export(temp_file, "PNG", width, height)

        # Read and encode as base64
        with open(temp_file, "rb") as f:
            image_data = f.read()

        base64_image = base64.b64encode(image_data).decode("utf-8")

        # Clean up temp file
        try:
            os.remove(temp_file)
        except:
            pass

        return {
            "success": True,
            "slide_number": slide_number,
            "width": width,
            "height": height,
            "format": "png",
            "base64_image": base64_image,
            "message": f"Screenshot of slide {slide_number} captured successfully"
        }

    except Exception as e:
        return {"error": f"Failed to capture screenshot: {str(e)}"}


@app.tool()
def screenshot_all_slides(
    width: int = 800,
    height: int = 450
) -> Dict:
    """
    Take screenshots of ALL slides in the presentation.
    Useful for getting an overview of the entire presentation.

    Args:
        width: Export width in pixels (smaller for overview)
        height: Export height in pixels

    Returns:
        Dict with list of base64 images and metadata
    """
    try:
        pres = get_active_presentation()
        slides_data = []

        for i in range(1, pres.Slides.Count + 1):
            slide = pres.Slides(i)
            temp_file = os.path.join(tempfile.gettempdir(), f"ppt_screenshot_{i}.png")
            slide.Export(temp_file, "PNG", width, height)

            with open(temp_file, "rb") as f:
                image_data = f.read()

            base64_image = base64.b64encode(image_data).decode("utf-8")

            slides_data.append({
                "slide_number": i,
                "base64_image": base64_image
            })

            try:
                os.remove(temp_file)
            except:
                pass

        return {
            "success": True,
            "total_slides": pres.Slides.Count,
            "width": width,
            "height": height,
            "format": "png",
            "slides": slides_data,
            "message": f"Captured {pres.Slides.Count} slide screenshots"
        }

    except Exception as e:
        return {"error": f"Failed to capture screenshots: {str(e)}"}


@app.tool()
def screenshot_window() -> Dict:
    """
    Take a screenshot of the entire PowerPoint window.
    Useful for seeing the PowerPoint UI state.

    Returns:
        Dict with base64 image data
    """
    try:
        import win32gui
        import win32ui
        import win32con

        ppt = get_ppt_app()

        # Find PowerPoint window
        hwnd = win32gui.FindWindow("PPTFrameClass", None)
        if not hwnd:
            # Try alternative class name
            hwnd = win32gui.FindWindow(None, ppt.ActiveWindow.Caption if ppt.Windows.Count > 0 else "PowerPoint")

        if not hwnd:
            return {"error": "Could not find PowerPoint window"}

        # Get window dimensions
        left, top, right, bottom = win32gui.GetWindowRect(hwnd)
        width = right - left
        height = bottom - top

        # Create device contexts
        hwndDC = win32gui.GetWindowDC(hwnd)
        mfcDC = win32ui.CreateDCFromHandle(hwndDC)
        saveDC = mfcDC.CreateCompatibleDC()

        # Create bitmap
        saveBitMap = win32ui.CreateBitmap()
        saveBitMap.CreateCompatibleBitmap(mfcDC, width, height)
        saveDC.SelectObject(saveBitMap)

        # Copy window content
        saveDC.BitBlt((0, 0), (width, height), mfcDC, (0, 0), win32con.SRCCOPY)

        # Convert to PIL Image
        bmpinfo = saveBitMap.GetInfo()
        bmpstr = saveBitMap.GetBitmapBits(True)

        img = Image.frombuffer(
            'RGB',
            (bmpinfo['bmWidth'], bmpinfo['bmHeight']),
            bmpstr, 'raw', 'BGRX', 0, 1
        )

        # Clean up
        win32gui.DeleteObject(saveBitMap.GetHandle())
        saveDC.DeleteDC()
        mfcDC.DeleteDC()
        win32gui.ReleaseDC(hwnd, hwndDC)

        # Convert to base64
        buffer = io.BytesIO()
        img.save(buffer, format="PNG")
        base64_image = base64.b64encode(buffer.getvalue()).decode("utf-8")

        return {
            "success": True,
            "width": width,
            "height": height,
            "format": "png",
            "base64_image": base64_image,
            "message": "PowerPoint window screenshot captured"
        }

    except Exception as e:
        return {"error": f"Failed to capture window screenshot: {str(e)}"}


# ============================================================================
# PRESENTATION MANAGEMENT
# ============================================================================

@app.tool()
def initialize_powerpoint() -> Dict:
    """
    Initialize PowerPoint application and make it visible.
    Call this first to ensure PowerPoint is running.

    Returns:
        Dict with status information
    """
    try:
        ppt = get_ppt_app()
        ppt.Visible = True

        return {
            "success": True,
            "message": "PowerPoint initialized and visible",
            "presentations_open": ppt.Presentations.Count,
            "version": ppt.Version
        }
    except Exception as e:
        return {"error": f"Failed to initialize PowerPoint: {str(e)}"}


@app.tool()
def create_presentation() -> Dict:
    """
    Create a new blank PowerPoint presentation.
    The presentation will be visible immediately in PowerPoint.

    Returns:
        Dict with presentation info
    """
    try:
        ppt = get_ppt_app()
        pres = ppt.Presentations.Add()

        return {
            "success": True,
            "message": "New presentation created",
            "name": pres.Name,
            "slide_count": pres.Slides.Count,
            "path": pres.FullName if pres.FullName else "(unsaved)"
        }
    except Exception as e:
        return {"error": f"Failed to create presentation: {str(e)}"}


@app.tool()
def open_presentation(file_path: str) -> Dict:
    """
    Open an existing PowerPoint presentation.

    Args:
        file_path: Full path to the .pptx file

    Returns:
        Dict with presentation info
    """
    try:
        if not os.path.exists(file_path):
            return {"error": f"File not found: {file_path}"}

        ppt = get_ppt_app()
        pres = ppt.Presentations.Open(file_path)

        return {
            "success": True,
            "message": f"Opened presentation: {pres.Name}",
            "name": pres.Name,
            "slide_count": pres.Slides.Count,
            "path": pres.FullName
        }
    except Exception as e:
        return {"error": f"Failed to open presentation: {str(e)}"}


@app.tool()
def save_presentation(file_path: Optional[str] = None) -> Dict:
    """
    Save the active presentation.

    Args:
        file_path: Path to save to (optional, uses current path if not specified)

    Returns:
        Dict with save status
    """
    try:
        pres = get_active_presentation()

        if file_path:
            pres.SaveAs(file_path)
            saved_path = file_path
        else:
            pres.Save()
            saved_path = pres.FullName

        return {
            "success": True,
            "message": f"Presentation saved",
            "path": saved_path
        }
    except Exception as e:
        return {"error": f"Failed to save presentation: {str(e)}"}


@app.tool()
def close_presentation(save: bool = True) -> Dict:
    """
    Close the active presentation.

    Args:
        save: Whether to save before closing

    Returns:
        Dict with status
    """
    try:
        pres = get_active_presentation()
        name = pres.Name

        if save:
            try:
                pres.Save()
            except:
                pass  # May fail if never saved

        pres.Close()

        return {
            "success": True,
            "message": f"Closed presentation: {name}"
        }
    except Exception as e:
        return {"error": f"Failed to close presentation: {str(e)}"}


@app.tool()
def get_presentation_info() -> Dict:
    """
    Get information about the active presentation.

    Returns:
        Dict with comprehensive presentation info
    """
    try:
        pres = get_active_presentation()

        # Get slide info
        slides_info = []
        for i in range(1, pres.Slides.Count + 1):
            slide = pres.Slides(i)
            slides_info.append({
                "number": i,
                "layout": slide.Layout,
                "shapes_count": slide.Shapes.Count,
                "name": slide.Name
            })

        return {
            "success": True,
            "name": pres.Name,
            "path": pres.FullName if pres.FullName else "(unsaved)",
            "slide_count": pres.Slides.Count,
            "slides": slides_info,
            "page_setup": {
                "width": pres.PageSetup.SlideWidth,
                "height": pres.PageSetup.SlideHeight
            }
        }
    except Exception as e:
        return {"error": f"Failed to get presentation info: {str(e)}"}


@app.tool()
def list_presentations() -> Dict:
    """
    List all open presentations in PowerPoint.

    Returns:
        Dict with list of open presentations
    """
    try:
        ppt = get_ppt_app()

        presentations = []
        for i in range(1, ppt.Presentations.Count + 1):
            pres = ppt.Presentations(i)
            presentations.append({
                "index": i,
                "name": pres.Name,
                "path": pres.FullName if pres.FullName else "(unsaved)",
                "slide_count": pres.Slides.Count
            })

        return {
            "success": True,
            "count": ppt.Presentations.Count,
            "presentations": presentations
        }
    except Exception as e:
        return {"error": f"Failed to list presentations: {str(e)}"}


# ============================================================================
# SLIDE MANAGEMENT
# ============================================================================

@app.tool()
def add_slide(
    layout: int = 2,
    position: Optional[int] = None
) -> Dict:
    """
    Add a new slide to the presentation.

    Layout types:
        1 = Title Slide
        2 = Title and Content (default)
        3 = Section Header
        4 = Two Content
        5 = Comparison
        6 = Title Only
        7 = Blank
        11 = Title and Vertical Text
        12 = Vertical Title and Text

    Args:
        layout: Slide layout type (1-12)
        position: Position to insert (None = end)

    Returns:
        Dict with new slide info
    """
    try:
        pres = get_active_presentation()

        if position is None:
            position = pres.Slides.Count + 1

        slide = pres.Slides.Add(position, layout)

        return {
            "success": True,
            "message": f"Added slide at position {position}",
            "slide_number": slide.SlideNumber,
            "layout": layout,
            "shapes_count": slide.Shapes.Count
        }
    except Exception as e:
        return {"error": f"Failed to add slide: {str(e)}"}


@app.tool()
def delete_slide(slide_number: int) -> Dict:
    """
    Delete a slide from the presentation.

    Args:
        slide_number: Slide number to delete (1-based)

    Returns:
        Dict with status
    """
    try:
        pres = get_active_presentation()

        if slide_number < 1 or slide_number > pres.Slides.Count:
            return {"error": f"Invalid slide number: {slide_number}"}

        pres.Slides(slide_number).Delete()

        return {
            "success": True,
            "message": f"Deleted slide {slide_number}",
            "remaining_slides": pres.Slides.Count
        }
    except Exception as e:
        return {"error": f"Failed to delete slide: {str(e)}"}


@app.tool()
def duplicate_slide(slide_number: int) -> Dict:
    """
    Duplicate a slide.

    Args:
        slide_number: Slide number to duplicate (1-based)

    Returns:
        Dict with new slide info
    """
    try:
        pres = get_active_presentation()

        if slide_number < 1 or slide_number > pres.Slides.Count:
            return {"error": f"Invalid slide number: {slide_number}"}

        slide = pres.Slides(slide_number)
        new_slide = slide.Duplicate()

        return {
            "success": True,
            "message": f"Duplicated slide {slide_number}",
            "new_slide_number": new_slide.SlideNumber
        }
    except Exception as e:
        return {"error": f"Failed to duplicate slide: {str(e)}"}


@app.tool()
def get_slide_info(slide_number: int) -> Dict:
    """
    Get detailed information about a specific slide.

    Args:
        slide_number: Slide number (1-based)

    Returns:
        Dict with slide details including all shapes
    """
    try:
        pres = get_active_presentation()

        if slide_number < 1 or slide_number > pres.Slides.Count:
            return {"error": f"Invalid slide number: {slide_number}"}

        slide = pres.Slides(slide_number)

        shapes_info = []
        for i in range(1, slide.Shapes.Count + 1):
            shape = slide.Shapes(i)
            shape_info = {
                "index": i,
                "name": shape.Name,
                "type": shape.Type,
                "left": round(shape.Left, 2),
                "top": round(shape.Top, 2),
                "width": round(shape.Width, 2),
                "height": round(shape.Height, 2)
            }

            # Add text if shape has text frame
            if shape.HasTextFrame:
                try:
                    shape_info["text"] = shape.TextFrame.TextRange.Text[:200]  # First 200 chars
                except:
                    pass

            shapes_info.append(shape_info)

        return {
            "success": True,
            "slide_number": slide_number,
            "name": slide.Name,
            "layout": slide.Layout,
            "shapes_count": slide.Shapes.Count,
            "shapes": shapes_info
        }
    except Exception as e:
        return {"error": f"Failed to get slide info: {str(e)}"}


@app.tool()
def go_to_slide(slide_number: int) -> Dict:
    """
    Navigate to a specific slide in the PowerPoint window.

    Args:
        slide_number: Slide number to navigate to (1-based)

    Returns:
        Dict with status
    """
    try:
        pres = get_active_presentation()
        ppt = get_ppt_app()

        if slide_number < 1 or slide_number > pres.Slides.Count:
            return {"error": f"Invalid slide number: {slide_number}"}

        # Navigate to slide
        ppt.ActiveWindow.View.GotoSlide(slide_number)

        return {
            "success": True,
            "message": f"Navigated to slide {slide_number}"
        }
    except Exception as e:
        return {"error": f"Failed to navigate: {str(e)}"}


# ============================================================================
# TEXT OPERATIONS
# ============================================================================

@app.tool()
def set_slide_title(
    slide_number: int,
    title: str
) -> Dict:
    """
    Set the title of a slide.

    Args:
        slide_number: Slide number (1-based)
        title: Title text

    Returns:
        Dict with status
    """
    try:
        pres = get_active_presentation()

        if slide_number < 1 or slide_number > pres.Slides.Count:
            return {"error": f"Invalid slide number: {slide_number}"}

        slide = pres.Slides(slide_number)

        # Find title shape
        for i in range(1, slide.Shapes.Count + 1):
            shape = slide.Shapes(i)
            if shape.HasTextFrame and shape.Type == 14:  # msoPlaceholder
                if shape.PlaceholderFormat.Type == 1:  # ppPlaceholderTitle
                    shape.TextFrame.TextRange.Text = title
                    return {
                        "success": True,
                        "message": f"Set title on slide {slide_number}",
                        "title": title
                    }

        return {"error": "No title placeholder found on this slide"}
    except Exception as e:
        return {"error": f"Failed to set title: {str(e)}"}


@app.tool()
def add_textbox(
    slide_number: int,
    text: str,
    left: float,
    top: float,
    width: float,
    height: float,
    font_size: Optional[int] = None,
    font_name: Optional[str] = None,
    font_color: Optional[str] = None,
    bold: bool = False,
    italic: bool = False,
    alignment: str = "left"
) -> Dict:
    """
    Add a text box to a slide.

    Args:
        slide_number: Slide number (1-based)
        text: Text content
        left: Left position in points
        top: Top position in points
        width: Width in points
        height: Height in points
        font_size: Font size in points (optional)
        font_name: Font name (optional)
        font_color: Hex color like "#FF0000" (optional)
        bold: Make text bold
        italic: Make text italic
        alignment: "left", "center", or "right"

    Returns:
        Dict with textbox info
    """
    try:
        pres = get_active_presentation()

        if slide_number < 1 or slide_number > pres.Slides.Count:
            return {"error": f"Invalid slide number: {slide_number}"}

        slide = pres.Slides(slide_number)

        # Add textbox (msoTextBox = 17)
        textbox = slide.Shapes.AddTextbox(1, left, top, width, height)
        textbox.TextFrame.TextRange.Text = text

        # Apply formatting
        text_range = textbox.TextFrame.TextRange

        if font_size:
            text_range.Font.Size = font_size
        if font_name:
            text_range.Font.Name = font_name
        if bold:
            text_range.Font.Bold = True
        if italic:
            text_range.Font.Italic = True

        if font_color:
            # Convert hex to RGB
            color = font_color.lstrip('#')
            r, g, b = int(color[0:2], 16), int(color[2:4], 16), int(color[4:6], 16)
            text_range.Font.Color.RGB = r + (g * 256) + (b * 256 * 256)

        # Set alignment
        align_map = {"left": 1, "center": 2, "right": 3}  # ppAlign values
        if alignment in align_map:
            text_range.ParagraphFormat.Alignment = align_map[alignment]

        return {
            "success": True,
            "message": f"Added textbox to slide {slide_number}",
            "shape_name": textbox.Name,
            "shape_index": textbox.ZOrderPosition
        }
    except Exception as e:
        return {"error": f"Failed to add textbox: {str(e)}"}


@app.tool()
def update_shape_text(
    slide_number: int,
    shape_index: int,
    text: str
) -> Dict:
    """
    Update text in an existing shape.

    Args:
        slide_number: Slide number (1-based)
        shape_index: Shape index (1-based)
        text: New text content

    Returns:
        Dict with status
    """
    try:
        pres = get_active_presentation()

        if slide_number < 1 or slide_number > pres.Slides.Count:
            return {"error": f"Invalid slide number: {slide_number}"}

        slide = pres.Slides(slide_number)

        if shape_index < 1 or shape_index > slide.Shapes.Count:
            return {"error": f"Invalid shape index: {shape_index}"}

        shape = slide.Shapes(shape_index)

        if not shape.HasTextFrame:
            return {"error": "Shape does not have a text frame"}

        shape.TextFrame.TextRange.Text = text

        return {
            "success": True,
            "message": f"Updated text in shape {shape_index} on slide {slide_number}",
            "shape_name": shape.Name
        }
    except Exception as e:
        return {"error": f"Failed to update text: {str(e)}"}


# ============================================================================
# SHAPE OPERATIONS
# ============================================================================

@app.tool()
def add_shape(
    slide_number: int,
    shape_type: str,
    left: float,
    top: float,
    width: float,
    height: float,
    fill_color: Optional[str] = None,
    line_color: Optional[str] = None,
    text: Optional[str] = None
) -> Dict:
    """
    Add a shape to a slide.

    Shape types: rectangle, rounded_rectangle, oval, triangle, diamond,
                 pentagon, hexagon, star, arrow_right, arrow_left,
                 arrow_up, arrow_down, heart, lightning_bolt

    Args:
        slide_number: Slide number (1-based)
        shape_type: Type of shape
        left: Left position in points
        top: Top position in points
        width: Width in points
        height: Height in points
        fill_color: Hex color like "#FF0000" (optional)
        line_color: Hex color for outline (optional)
        text: Text to add inside shape (optional)

    Returns:
        Dict with shape info
    """
    try:
        pres = get_active_presentation()

        if slide_number < 1 or slide_number > pres.Slides.Count:
            return {"error": f"Invalid slide number: {slide_number}"}

        slide = pres.Slides(slide_number)

        # Shape type mapping (msoAutoShapeType values)
        shape_map = {
            "rectangle": 1,
            "rounded_rectangle": 5,
            "oval": 9,
            "triangle": 7,
            "diamond": 4,
            "pentagon": 51,
            "hexagon": 10,
            "star": 12,
            "arrow_right": 33,
            "arrow_left": 34,
            "arrow_up": 35,
            "arrow_down": 36,
            "heart": 21,
            "lightning_bolt": 22
        }

        shape_type_lower = shape_type.lower()
        if shape_type_lower not in shape_map:
            return {
                "error": f"Unknown shape type: {shape_type}",
                "available_types": list(shape_map.keys())
            }

        shape = slide.Shapes.AddShape(shape_map[shape_type_lower], left, top, width, height)

        # Apply fill color
        if fill_color:
            color = fill_color.lstrip('#')
            r, g, b = int(color[0:2], 16), int(color[2:4], 16), int(color[4:6], 16)
            shape.Fill.Solid()
            shape.Fill.ForeColor.RGB = r + (g * 256) + (b * 256 * 256)

        # Apply line color
        if line_color:
            color = line_color.lstrip('#')
            r, g, b = int(color[0:2], 16), int(color[2:4], 16), int(color[4:6], 16)
            shape.Line.ForeColor.RGB = r + (g * 256) + (b * 256 * 256)

        # Add text
        if text:
            shape.TextFrame.TextRange.Text = text

        return {
            "success": True,
            "message": f"Added {shape_type} to slide {slide_number}",
            "shape_name": shape.Name,
            "shape_index": shape.ZOrderPosition
        }
    except Exception as e:
        return {"error": f"Failed to add shape: {str(e)}"}


@app.tool()
def add_image(
    slide_number: int,
    image_path: str,
    left: float,
    top: float,
    width: Optional[float] = None,
    height: Optional[float] = None
) -> Dict:
    """
    Add an image to a slide.

    Args:
        slide_number: Slide number (1-based)
        image_path: Full path to image file
        left: Left position in points
        top: Top position in points
        width: Width in points (optional, maintains aspect ratio if only one specified)
        height: Height in points (optional)

    Returns:
        Dict with image shape info
    """
    try:
        pres = get_active_presentation()

        if slide_number < 1 or slide_number > pres.Slides.Count:
            return {"error": f"Invalid slide number: {slide_number}"}

        if not os.path.exists(image_path):
            return {"error": f"Image file not found: {image_path}"}

        slide = pres.Slides(slide_number)

        # Add picture
        if width and height:
            picture = slide.Shapes.AddPicture(
                image_path, False, True, left, top, width, height
            )
        else:
            picture = slide.Shapes.AddPicture(
                image_path, False, True, left, top
            )
            if width:
                ratio = width / picture.Width
                picture.Width = width
                picture.Height = picture.Height * ratio
            elif height:
                ratio = height / picture.Height
                picture.Height = height
                picture.Width = picture.Width * ratio

        return {
            "success": True,
            "message": f"Added image to slide {slide_number}",
            "shape_name": picture.Name,
            "width": round(picture.Width, 2),
            "height": round(picture.Height, 2)
        }
    except Exception as e:
        return {"error": f"Failed to add image: {str(e)}"}


@app.tool()
def delete_shape(
    slide_number: int,
    shape_index: int
) -> Dict:
    """
    Delete a shape from a slide.

    Args:
        slide_number: Slide number (1-based)
        shape_index: Shape index (1-based)

    Returns:
        Dict with status
    """
    try:
        pres = get_active_presentation()

        if slide_number < 1 or slide_number > pres.Slides.Count:
            return {"error": f"Invalid slide number: {slide_number}"}

        slide = pres.Slides(slide_number)

        if shape_index < 1 or shape_index > slide.Shapes.Count:
            return {"error": f"Invalid shape index: {shape_index}"}

        shape_name = slide.Shapes(shape_index).Name
        slide.Shapes(shape_index).Delete()

        return {
            "success": True,
            "message": f"Deleted shape '{shape_name}' from slide {slide_number}"
        }
    except Exception as e:
        return {"error": f"Failed to delete shape: {str(e)}"}


# ============================================================================
# TABLE OPERATIONS
# ============================================================================

@app.tool()
def add_table(
    slide_number: int,
    rows: int,
    cols: int,
    left: float,
    top: float,
    width: float,
    height: float,
    data: Optional[List[List[str]]] = None
) -> Dict:
    """
    Add a table to a slide.

    Args:
        slide_number: Slide number (1-based)
        rows: Number of rows
        cols: Number of columns
        left: Left position in points
        top: Top position in points
        width: Width in points
        height: Height in points
        data: Optional 2D list of cell values

    Returns:
        Dict with table info
    """
    try:
        pres = get_active_presentation()

        if slide_number < 1 or slide_number > pres.Slides.Count:
            return {"error": f"Invalid slide number: {slide_number}"}

        slide = pres.Slides(slide_number)

        # Add table
        table_shape = slide.Shapes.AddTable(rows, cols, left, top, width, height)
        table = table_shape.Table

        # Populate data if provided
        if data:
            for row_idx, row_data in enumerate(data):
                if row_idx >= rows:
                    break
                for col_idx, cell_text in enumerate(row_data):
                    if col_idx >= cols:
                        break
                    table.Cell(row_idx + 1, col_idx + 1).Shape.TextFrame.TextRange.Text = str(cell_text)

        return {
            "success": True,
            "message": f"Added {rows}x{cols} table to slide {slide_number}",
            "shape_name": table_shape.Name,
            "shape_index": table_shape.ZOrderPosition
        }
    except Exception as e:
        return {"error": f"Failed to add table: {str(e)}"}


@app.tool()
def set_table_cell(
    slide_number: int,
    shape_index: int,
    row: int,
    col: int,
    text: str
) -> Dict:
    """
    Set text in a table cell.

    Args:
        slide_number: Slide number (1-based)
        shape_index: Table shape index (1-based)
        row: Row index (1-based)
        col: Column index (1-based)
        text: Cell text

    Returns:
        Dict with status
    """
    try:
        pres = get_active_presentation()

        if slide_number < 1 or slide_number > pres.Slides.Count:
            return {"error": f"Invalid slide number: {slide_number}"}

        slide = pres.Slides(slide_number)

        if shape_index < 1 or shape_index > slide.Shapes.Count:
            return {"error": f"Invalid shape index: {shape_index}"}

        shape = slide.Shapes(shape_index)

        if not shape.HasTable:
            return {"error": "Shape is not a table"}

        table = shape.Table

        if row < 1 or row > table.Rows.Count:
            return {"error": f"Invalid row: {row}"}
        if col < 1 or col > table.Columns.Count:
            return {"error": f"Invalid column: {col}"}

        table.Cell(row, col).Shape.TextFrame.TextRange.Text = text

        return {
            "success": True,
            "message": f"Set cell ({row}, {col}) text"
        }
    except Exception as e:
        return {"error": f"Failed to set cell: {str(e)}"}


# ============================================================================
# MAIN EXECUTION
# ============================================================================

def main():
    """Run the FastMCP server."""
    app.run(transport='stdio')


if __name__ == "__main__":
    main()
