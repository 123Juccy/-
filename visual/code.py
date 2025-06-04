import io
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from PIL import Image, ImageTk
import barcode
from barcode.writer import ImageWriter
import qrcode
from openpyxl import load_workbook
from datetime import datetime
import os
import threading
import time
import win32print
import win32api
import tempfile
import pywintypes
import win32con
import win32ui
from PIL import ImageWin
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont