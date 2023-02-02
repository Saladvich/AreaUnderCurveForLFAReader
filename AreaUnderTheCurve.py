#Gupta, R.; Gupta, P.; Wang, S.; Melnykov, A.; Jiang, Q.; Seth, A.; Wang, Z.; Morrissey, J.; George, I.; Gandra, S.; Sinha, P.; Storch, G.; Parikh, B.; Genin, G.; Singamanenei, S.; 
#Ultrasensitive lateral-flow assays via plasmonically active antibody-conjugated fluorescent nanoparticles. 
#Nature Biomedical Engineering
#DOI: https://www.nature.com/articles/s41551-022-01001-1

import numpy as np
import matplotlib.pyplot as plt
from openpyxl import load_workbook
from scipy.ndimage import uniform_filter1d
from pybaselines.spline import mixture_model
from pybaselines.polynomial import imodpoly
from scipy.signal import find_peaks


path = "" #Path of excel File
wb = load_workbook(path)
sheet = wb[""] #Sheet of excel data
columns = [""] #columns to read
printLine = "" #Excel row under the read column where calculation will be printed
areaRange = "" #Range for area calculation ex. 25 would mean ±25 for a total of 50mm of integration


def find_nearest(array, value):
    array = np.asarray(array)
    idx = (np.abs(array - value)).argmin()
    return array[idx]

def AUC(x1, x2, arr):
    area = 0
    x1 = np.where(x == (find_nearest(x, x1)))
    x2 = np.where(x == (find_nearest(x, x2)))

    x1 = int(x1[0])
    x2 = int(x2[0])

    for i in range(x1, x2 + 1, 1):
        y1 = arr[i]
        y2 = arr[i + 1]

        trapArea = ((y1 + y2)/2)
        area += trapArea

    return area

def findCenter(fit):
    peaks = find_peaks(fit, threshold=0, height=0.1, distance=15)
    peak_pos = x[peaks[0]]
    if fit.max() > 8:
        highBound = 140
        lowBound = 80
    else:
        highBound = 125
        lowBound = 80

    for pos in peak_pos:
        if pos >= highBound or pos <= lowBound:
            peak_pos = np.delete(peak_pos, np.where(peak_pos == pos))
    height = fit[peak_pos.astype(int)]
    max_index = peak_pos[height.argmax()]

    y2 = fit * -1
    minima = find_peaks(y2, height=-0.1, distance=15)
    min_pos = x[minima[0]]
    for pos in min_pos:
        if pos >= 156 or pos <= 75:
            min_pos = np.delete(min_pos, np.where(min_pos == pos))

    if min_pos[len(min_pos) - 1] <= max_index:
        min_pos = np.append(min_pos, max_index + 20)
    elif min_pos[0] >= max_index:
        min_pos = np.append(min_pos, max_index - 20)
    min_height = fit[min_pos.astype(int)]

    focals = np.array([])
    focalCoords = np.array([])
    for i in range(0, len(min_pos) - 1, 1):
        focals = np.append(focals, int((min_pos[i] + min_pos[i + 1] + max_index) / 3))
        focalCoords = np.append(focalCoords, str(min_pos[i]) + " " + str(min_pos[i + 1]) + " " + str(max_index))

    weighted_focals = np.abs(focals - max_index)
    focal = focals[weighted_focals.argmin()]
    triangulation = focalCoords[weighted_focals.argmin()].split()

    coordsT = np.array([])
    coordsTx = np.array([])
    for i in triangulation:
        coordT = int(float(i))
        coordsTx = np.append(coordsTx, coordT)
        coordsT = np.append(coordsT, fit[coordT])

    centroid_y = (coordsT[0] + coordsT[1] + coordsT[2]) / 3

    return [min_height, focal, centroid_y]


if __name__ == "__main__":
    control_line = False
    for col in columns:
        try:
            if isinstance(sheet[col + printLine].value, float) != True:
                column = sheet[col]
                y_original = np.array([])
                for cell in column:
                    if isinstance(cell.value, float):
                        y_original = np.append(y_original, cell.value)
                if len(y_original) > 20:
                    print(col)
                    lowCut = 23
                    highCut = len(y_original) - 1
                    y = y_original[lowCut:highCut]

                    smooth = uniform_filter1d(y, 5)

                    poly = imodpoly(smooth, poly_order=2, max_iter=50, num_std=1)[0]
                    fit = y - poly

                    if fit[130 - lowCut:len(fit)].max() > ((fit[70 - lowCut:125 - lowCut].max())*1.5) or control_line == True:
                        highCut = 125
                        y = y_original[lowCut:highCut]
                    smooth = uniform_filter1d(y, 5)
                    poly = imodpoly(smooth, poly_order=2, max_iter=50, num_std=1)[0]
                    fit = y - poly
                    mixturePoly, params_1 = mixture_model(fit, lam=1e6, diff_order=3, max_iter=100)
                    fit = fit - mixturePoly

                    for i in range(0, lowCut, 1):
                        fit = np.insert(fit, 0, 0)
                    for i in range(highCut, len(y_original) - 1, 1):
                        fit = np.insert(fit, i, 0)
                    x = np.linspace(0, len(fit), len(fit))

                    if fit[145:160].max() > 15 or control_line == False:
                        # fig = plt.figure()
                        # plt.plot(y_original, color='black', label='Original', fontsize=13, lw=2.3)
                        # plt.xlabel("Distance from conjugate pad (0.1 mm)", fontsize=13, weight='bold')
                        # plt.ylabel("Signal Intensity (a.u.)", fontsize=13, weight='bold')
                        # plt.tick_params(axis='x', labelsize=11)
                        # plt.tick_params(axis='y', labelsize=11)
                        # plt.legend(fontsize=14)
                        # plt.show()
                        
                        try:
                            variables = findCenter(fit)
                            min_height = variables[0]
                            focal = variables[1]
                            centroid_y = variables[2]
                            # plt.plot(x, fit, color='red', label='Corrected', fontsize=13, lw=2.3)
                            # plt.fill_between(x, fit, where=(x <= focal + areaRange) & (x >= focal - areaRange), color="red", alpha=0.2)
                            # plt.scatter(peak_pos, height, color='indigo', s=50, marker='D', label='Local Maxima')
                            # plt.scatter(min_pos, min_height, color='blue', s=50, marker='D', label='Local Minima')
                            # plt.scatter(coordsTx, coordsT, color='green', s=55, marker='X', label='Triangle Verticies')
                            # plt.scatter(focal, centroid_y, color='green', s=40, marker='o', label='Centroid')
                            # plt.plot([focal - areaRange, focal + areaRange], [centroid_y, centroid_y], marker='o', color='green', lw=2, label='Range of Integration (±'+areaRange+')')
                            # plt.xlabel("Distance from conjugate pad (0.1 mm)", fontsize=13, weight='bold')
                            # plt.ylabel("Signal Intensity (a.u.)", fontsize=13, weight='bold')
                            # plt.tick_params(axis='x', labelsize=11)
                            # plt.tick_params(axis='y', labelsize=11)
                            # plt.legend(fontsize=14)
                            # plt.show()
                        except TypeError:
                            pass

                        sheet[col + printLine] = (round(AUC(focal - areaRange, focal + areaRange, fit), 4))
                        print(round(AUC(focal - areaRange, focal + areaRange, fit), 4))
                        wb.save(path)
                    else:
                        print("ERROR: Invalid Test. Control Line Not Present")
        except TypeError:
            pass
