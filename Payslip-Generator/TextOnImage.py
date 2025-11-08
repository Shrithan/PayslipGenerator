from PIL import Image, ImageDraw, ImageFont

def generate_payslip_overlay(
    template_path,
    save_path,
    month, name, designation, department, calendar_days, ctc,
    emp_code, email, lop, present_days,
    bank_details,
    basic, hra, conveyance, special_allowance,
    arrears, bonus, gross, balance_payable,
    provident_fund, employee_state_insurance, pt, tds, deduction, arrears_withheld, total_deduction
):
    image = Image.open(template_path)
    draw = ImageDraw.Draw(image)
    font_path_regular = "Tinos-Regular.ttf"
    font_path_bold = "Tinos-Bold.ttf"

    def draw_text(text, position, bold=False, align_right=False):
        font = ImageFont.truetype(font_path_bold if bold else font_path_regular, 20)
        if align_right:
            text_width = font.getbbox(text)[2]
            position = (position[0] - text_width, position[1])
        draw.text(position, text, fill=(0, 0, 0), font=font)

    # ************ Left Block ************
    draw_text(month, (650, 126), bold=True)
    draw_text(name, (245, 183))
    draw_text(designation, (245, 212))
    draw_text(department, (245, 241))
    draw_text(str(calendar_days), (245, 267))
    draw_text(str(ctc), (245, 295))

    # ************ Right Block ************
    draw_text(emp_code, (721, 183))
    draw_text(email, (721, 212))
    draw_text(str(present_days), (721, 241))
    draw_text(str(lop), (721, 267))
    draw_text(bank_details, (721, 295))

    # ************ Earnings (Left Table) ************
    draw_text(str(basic), (515, 376), align_right=True)
    draw_text(str(hra), (515, 406), align_right=True)
    draw_text(str(conveyance), (515, 434), align_right=True)
    draw_text(str(special_allowance), (515, 462), align_right=True)
    draw_text(str(bonus), (515, 487), align_right=True)
    draw_text(str(arrears), (515, 516), align_right=True)
    draw_text(str(gross), (515, 547), bold=True, align_right=True)
    draw_text(str(balance_payable), (515, 575), bold=True, align_right=True)

    # ************ Deductions (Right Table) ************
    draw_text(str(provident_fund), (1050, 376), align_right=True)
    draw_text(str(employee_state_insurance), (1050, 406), align_right=True)
    draw_text(str(pt), (1050, 434), align_right=True)
    draw_text(str(tds), (1050, 462), align_right=True)
    draw_text(str(deduction), (1050, 487), align_right=True)
    draw_text(str(arrears_withheld), (1050, 516), align_right=True)
    draw_text(str(total_deduction), (1050, 547), bold=True, align_right=True)

    rgb_image = image.convert("RGB")
    rgb_image.save(save_path.replace(".jpg", ".pdf"))

