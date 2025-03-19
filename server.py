from flask import Flask, request, redirect, send_file
import logging
import io
from PIL import Image
import datetime

app = Flask(__name__)

# Configure logging to output to a file named "tracking.log"
logging.basicConfig(filename="tracking.log", level=logging.INFO, 
                    format='%(asctime)s - %(levelname)s - %(message)s')

def get_transparent_pixel():
    """
    Create a 1x1 transparent PNG image and return it as a BytesIO object.
    """
    img = Image.new("RGBA", (1, 1), (0, 0, 0, 0))
    img_byte_arr = io.BytesIO()
    img.save(img_byte_arr, format="PNG")
    img_byte_arr.seek(0)
    return img_byte_arr

@app.route('/track_pixel')
def track_pixel():
    """
    Endpoint to serve a tracking pixel.
    Logs the email and timestamp when the pixel is loaded.
    """
    email = request.args.get('email', 'unknown')
    timestamp = datetime.datetime.now().isoformat()
    logging.info(f"Tracking Pixel Loaded - Email: {email} at {timestamp}")
    # Return a transparent 1x1 PNG image
    return send_file(get_transparent_pixel(), mimetype='image/png')

@app.route('/track_link')
def track_link():
    """
    Endpoint for tracking link clicks.
    Logs the click event with email and then redirects to the specified URL.
    """
    email = request.args.get('email', 'unknown')
    redirect_url = request.args.get('redirect', '/')
    timestamp = datetime.datetime.now().isoformat()
    logging.info(f"Link Clicked - Email: {email} at {timestamp}, Redirecting to: {redirect_url}")
    return redirect(redirect_url)

if __name__ == '__main__':
    # Run the server on port 5000 (or any port you prefer)
    app.run(port=5000, debug=True)
