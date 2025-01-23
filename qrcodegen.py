import qrcode
from qrcode.constants import ERROR_CORRECT_L

class QRcodeGen:
    def __init__(self):
        self.qr = qrcode.QRCode(
            version=1, 
            error_correction=ERROR_CORRECT_L, 
            box_size=10,
            border=4, 
        )

    def __call__(self,data:str, path:str) -> None:
        self.qr.add_data(data)
        self.qr.make(fit=True)
        img = self.qr.make_image(fill_color="black", back_color="white")
        img.save(path)
        
