import qrcode
from qrcode.constants import ERROR_CORRECT_L

class QRcodeGen:
    def __init__(self):
        # Initialize the QR code object with specified parameters
        self.qr = qrcode.QRCode(
            version=1,                 # Controls the size of the QR Code (version 1 is 21x21)
            error_correction=ERROR_CORRECT_L,  # Error correction level (L: approximately 7% or less errors corrected)
            box_size=10,              # Size of each box in the QR code
            border=4,                 # Border size, in boxes
        )

    def __call__(self, data: str, path: str) -> None:
        """
        Generate a QR code from the given data and save it to the specified path.

        :param data: The data to encode in the QR code.
        :param path: The file path where the QR code image will be saved.
        """
        self.qr.add_data(data)  # Add data to the QR code
        self.qr.make(fit=True)  # Compile the data into a QR code array (fit=True ensures the QR code size fits the data)

        # Create an image from the QR code array
        img = self.qr.make_image(fill_color="black", back_color="white")

        # Save the image to the specified path
        img.save(path)

# Example usage:
# qr_gen = QRcodeGen()
# qr_gen("https://www.example.com", "qrcode.png")
