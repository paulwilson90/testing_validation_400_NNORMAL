import json
import math
import pandas as pd

RED = '\033[31m'
REDEND = '\033[0m'


def get_uld(elevation, flap, weight):
    """Gets the ULD by interpolating and using index locations from the QRH
    It grabs the weight one tonne up and below and the elevation INDEX position one up and below.
    It then interpolates using the percentage of the remaining index location."""
    if flap == 0 or flap == 5:
        flap = 35
    weight_tonnes = weight / 1000
    flap = str(int(flap))
    wt_up = str(math.ceil(float(weight_tonnes)))
    wt_down = str(math.floor(float(weight_tonnes)))
    with open('ulds_q400.json') as ulds:
        uld_ = json.load(ulds)
    elevation_up = math.ceil(elevation)
    elevation_down = math.floor(elevation)
    # interpolating with the upper weight of the two elevation figures
    wt_up_up_data = uld_[flap][wt_up][elevation_up]
    wt_up_dwn_data = uld_[flap][wt_up][elevation_down]
    uld_up_wt = round(wt_up_dwn_data + ((wt_up_up_data - wt_up_dwn_data) * (elevation - elevation_down)))
    # interpolating with the lower weight of the two elevation figures
    wt_dwn_up_data = uld_[flap][wt_down][elevation_up]
    wt_dwn_dwn_data = uld_[flap][wt_down][elevation_down]
    uld_dwn_wt = round(wt_dwn_dwn_data + ((wt_dwn_up_data - wt_dwn_dwn_data) * (elevation - elevation_down)))
    # interpolating for weight between the two elevation interpolated figures
    final_uld = round(uld_dwn_wt + (uld_up_wt - uld_dwn_wt) * (float(weight_tonnes) - int(wt_down)))

    return final_uld


def wind_correct_formulated(ULD, wind_comp):
    """For every ULD entry to the wind chart above 700, add 0.003m on top of 3.8 for every knot head
    for every ULD entry to the wind chart above 700, add 0.01m on top of 12 for every knot tail"""
    amount_above_700 = ULD - 700
    if wind_comp > 0:  # if its a headwind
        factor_above_uld = amount_above_700 * 0.003
        wind_corrected_ULD = round(ULD - (wind_comp * (3.8 + factor_above_uld)))
    else:  # if its a tailwind
        factor_above_uld = amount_above_700 * 0.01
        wind_corrected_ULD = ULD - round((wind_comp * (12 + factor_above_uld)))

    if wind_comp < -10:  # if the wind is more than 10 knot tail, add 1.6% for every knot over 10t
        factor_above_uld = amount_above_700 * 0.01
        ten_tail_ULD = ULD - round((-10 * (12 + factor_above_uld)))
        wind_corrected_ULD = int(ten_tail_ULD * (1 + ((abs(wind_comp) - 10) * 1.6) / 100))

    return int(wind_corrected_ULD)


def slope_corrected(slope, wind_corrected_ld):
    """If the slope is greater than 0, the slope is going uphill so the distance will be shorter
    IF the slope is less than 0 however, the slope is downhill and the distance increases.
    For every 1% slope downhill (Negative slope), increase the ULD by 9.25% 630
    For every 1% slope uphill (Positive slope), decrease the ULD by 6.5%"""
    #  if the slope is downhill
    if slope < 0:
        slope_correct = wind_corrected_ld + (wind_corrected_ld * (abs(slope) * 0.0925))
    #  if the slope is uphill
    else:
        slope_correct = wind_corrected_ld - (wind_corrected_ld * (abs(slope) * 0.065))
    return slope_correct


def get_v_speeds(weight, flap, vapp_addit, ice, ab_fctr):
    """Using data from AFM 5-1-2 charts for appropriate flap setting, finding the VSR 1.23 speeds (VRef Speeds)
    Weight is rounded up to the nearest 500 to determine the VRef.
    For ice protection addits:
    flap 0 is 25
    flap 5, 10 and 15 are 20
    flap 35 is 15"""
    can_land_in_this_config = True
    flap = str(flap)
    weight = str((math.ceil(weight / 500) * 500) / 1000)
    print(f"Using {weight}t as the weight to get VREF")
    # reading the excel file
    xl = pd.ExcelFile('400_MELCDL_MULTIPLIERS.xlsx')
    Q400 = pd.read_excel(xl, 'NON NORMAL')
    # getting the appropriate speed addit or VS, if none apply, the speed variable will return nan...
    for line in range(len(Q400)):
        all_rows = Q400.loc[line]
        if all_rows['Problem'] == ab_fctr:
            speed = all_rows['F' + flap + " Add"]
    # get the unaltered 1.3 VREF speed
    with open('ref_speeds.json') as file:
        f = json.load(file)
    vref = f[flap][weight]
    # if the QRH specifies a speed for approach for the specific failure, determine whether its a 1.4vs
    # or an additive to the 1.3 VS/VREF and apply. This will become the new VREF.
    if not pd.isnull(speed):
        print("There is a QRH prescribed landing speed", speed)
        if speed == 1.3:
            with open("one_point_three.json") as one_point_four:
                o_p_f = json.load(one_point_four)
                vref = o_p_f[flap][weight]
        else:
            vref = int(vref + speed)
    # apply the INCR REF speed applicable to flap setting
    vapp = int(vref) + vapp_addit
    if ab_fctr == "LOSS OF ALL FLUID FROM NO.1 HYDRAULIC SYSTEM" or ab_fctr == "NO.1 AND NO.2 HYDAULIC SYSTEMS FAILURE":
        if flap == "0":
            vref_ice = vref + 25
        elif flap == "15" or flap == "10" or flap == "5":
            vref_ice = vref + 20
        else:
            vref_ice = vref + 25
    elif (ab_fctr == "DEICE PRESS") and (ice == "On"):
        if flap == "10":
            vref_ice = vref + 30
        elif flap == "15":
            vref_ice = vref + 25
        else:
            can_land_in_this_config = False
    elif ab_fctr == "ROLL SPLR INBD HYD OR ROLL SPLR OUTBD HYD (CAUTION LIGHT)" \
                    "" or ab_fctr == "LOSS OF ALL FLUID FROM NO.2 HYDRAULIC SYSTEM" \
                                     "" or ab_fctr == "#1 HYD ISO VLV (CAUTION LIGHT)" \
                                                      "" or ab_fctr == "#2 HYD ISO VLV (CAUTION LIGHT)":
        if flap == "10":
            vref_ice = vref + 20
        elif flap == "15":
            vref_ice = vref + 20
        elif flap == "35":
            vref_ice = vref + 25
        else:
            can_land_in_this_config = False
    else:
        if flap == "0":
            vref_ice = vref + 25
        elif flap == "15" or flap == "10" or flap == "5":
            vref_ice = vref + 20
        else:
            vref_ice = vref + 15

    # if the ice protection is ON, then VAPP will become VREF ICE
    if ice == "On":
        vapp = vref_ice
    print(vref, "VREF ADDIT", vapp_addit, "VAPP", vapp, "VREF ICE", vref_ice)

    return vapp, vref, vref_ice, can_land_in_this_config


def abnormal_factor(ab_fctr, corrected_for_slope, flap, ice):
    """Take in the abnormal factor from the excel sheet and pull its factor from the Multipliers excel sheet
    Return the landing distance required after applying the factor to the slope corrected distance. This is
    either the ice ON or OFF distance, not both....
    Return the multiplier used to get the distance.
    If N/A is listed in the MELCDL_MULTIPLIERS sheet for the current flap setting for the abnormal, a parameter
    can_land_in_this_config is returned as false and the remaining calculations won't be displayed in final sheet.
    """
    print(ab_fctr, "Is the Abnormality")
    can_land_in_this_config = True
    flap = str(flap)
    xl = pd.ExcelFile('400_MELCDL_MULTIPLIERS.xlsx')
    Q400 = pd.read_excel(xl, 'NON NORMAL')
    for line in range(len(Q400)):
        all_rows = Q400.loc[line]
        if all_rows['Problem'] == ab_fctr:
            multiplier = all_rows['F' + flap + " " + ice]
    if ab_fctr == "EXTENDED DOOR OPEN" or ab_fctr == "EXTENDED DOOR CLOSED":  # due to it being WAT and MLDW issue only
        multiplier = 1
    if pd.isnull(multiplier):  # means the multiplier is N/A and not for landing in this config
        multiplier = 1
        can_land_in_this_config = False

    distance = corrected_for_slope * multiplier

    print("Abnormal Multiplier with the ice protection", ice, "is", multiplier, "giving a new distance required of",
          distance)
    return int(distance), multiplier, can_land_in_this_config


def vapp_corrections(abnormal_dist, vref_addit, wet_dry):
    """Apply Operational Correction
    1.20 (Dry VREF+10)  # Every knot is 1.02
    1.50 (Wet VREF)
    1.70 (Wet VREF+10)  # Every knot is 1.02
    With wet, starting with 1.5 as the base and adding 0.02 for every knot additional"""

    if wet_dry == "Wet":
        percent_increase = 1.5 + (vref_addit * 0.02)
    else:
        percent_increase = 1 + (vref_addit * 0.02)

    abnormal_vapp_adjusted_ld = abnormal_dist * percent_increase

    print(f"It is {wet_dry} VREF + {vref_addit} so the multiplier is {percent_increase}."
          f"This gives us {int(abnormal_vapp_adjusted_ld)} as the distance")

    return int(abnormal_vapp_adjusted_ld)


def company_addit_dry_wet(vapp_corrected_ld):
    """Applying 15% addition to the vapp corrected landing distance"""
    operational_fact_adjusted_ld = vapp_corrected_ld * 1.15
    return int(operational_fact_adjusted_ld)


def get_wat_limit(temp, flap, propeller_rpm, bleed, pressure_alt, test_case):
    """Take in the temp, flap, bleed position and pressure altitude as parameters
    and return the max landing weight.
    Also trying to keep indexes in range as some temperatures and pressure altitudes are off charts.
    The minimum pressure alt for the chart is 0 and the max is 4000.
    The minimum temperature is 0 and the max is 48, even after the 11 degree addit"""
    off_chart_limits = False
    MLDW = 28009

    flap = str(int(flap))
    if pressure_alt < 0:
        pressure_alt = 0
        off_chart_limits = True
    else:
        if pressure_alt > 4000:
            pressure_alt = 4000 / 500
            off_chart_limits = True
        else:
            pressure_alt = pressure_alt / 500
    if propeller_rpm == "RDCP":
        rpm = "850"
    else:
        rpm = "1020"
    if bleed == "On":
        temp = int(temp) + 11

    if temp > 48:
        temp = str(48)
        off_chart_limits = True
        if pressure_alt > 2:
            pressure_alt = 2
    else:
        if temp < 0:
            temp = str(0)
            off_chart_limits = True
        else:
            temp = str(temp)
    if flap == "35":
        ga_flap = "15"
    else:
        ga_flap = "10"

    with open(f'wat_f{ga_flap}.json') as r:
        wat = json.load(r)
    elev_up = math.ceil(pressure_alt)
    elev_down = math.floor(pressure_alt)
    temp_up = str(math.ceil(int(temp) / 2) * 2)
    temp_down = str(math.floor(int(temp) / 2) * 2)

    # interpolating with the upper temp of the two elevation figures
    try:
        temp_up_up_data = wat[rpm][temp_up][elev_up]
    except Exception as err:
        print(RED + "ERROR" + REDEND, err, "TEST CASE", test_case)

    temp_up_dwn_data = wat[rpm][temp_up][elev_down]
    temp_up_wt = round(temp_up_dwn_data + ((temp_up_up_data - temp_up_dwn_data) * (pressure_alt - elev_down)))
    # interpolating with the lower temp of the two elevation figures
    temp_dwn_up_data = wat[rpm][temp_down][elev_up]
    temp_dwn_dwn_data = wat[rpm][temp_down][elev_down]
    temp_dwn_wt = round(temp_dwn_dwn_data + ((temp_dwn_up_data - temp_dwn_dwn_data) * (pressure_alt - elev_down)))

    wat_limit = int((temp_up_wt + temp_dwn_wt) / 2)

    if flap == "5" or flap == "0" or flap == "10":  # Should be able to climb with no WAT limit at these flap settings
        return 28009, MLDW, off_chart_limits

    return wat_limit, MLDW, off_chart_limits


def max_landing_wt_lda(lda, operation_fact_corrected_ld, flap, weight, unfact_uld):
    """Find the ratio between the landing distance required and the unfactored ULD which returns a multiplier ratio
    Divide the landing distance available by the ratio to find the relative unfactored ULD
    Get the difference between the maximum (LDA based) ULD and the current ULD and divide by 23 for flap 15 or 20.5 for
    flap 35 and multiply by 1000 (This is ULD difference for every tonne) this will give the weight to add onto the
    current landing weight which will give the max field landing weight."""
    flap = str(flap)
    if flap == "15":
        ratio = operation_fact_corrected_ld / unfact_uld
        max_unfact_uld = lda / ratio
        diff_between_ulds = max_unfact_uld - unfact_uld
        final = ((diff_between_ulds / 23) * 1000) + weight
    else:
        ratio = operation_fact_corrected_ld / unfact_uld
        max_unfact_uld = lda / ratio
        diff_between_ulds = max_unfact_uld - unfact_uld
        final = ((diff_between_ulds / 20.5) * 1000) + weight
    return int(final)


def final_max_weight(max_wat, max_field, MLDW, off_chart):
    """Find and return the lowest weight out of all provided. Also add * to any code where the wat weight
    used a parameter that was off chart."""
    # f means field, s means struc, c means climb
    if max_wat < max_field:
        max_weight = max_wat
        code_max = "(c)"
    else:
        max_weight = max_field
        code_max = "(f)"
    if max_weight > MLDW:
        max_weight = MLDW
        code_max = "(s)"

    if off_chart:
        max_weight = str(max_weight) + code_max + "^"
    else:
        max_weight = str(max_weight) + code_max
    return max_weight
