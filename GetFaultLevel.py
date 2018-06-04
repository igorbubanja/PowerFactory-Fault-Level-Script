'''
Python scripts to get fault data
NOTE: The Trasnpower system uses Python 3.5
'''
# Powerfactory stuff
import powerfactory as pf
import numpy as np
import csv

app = pf.GetApplication()
if app is None:
	raise Exception("Cannot get Powerfactory link")
	
def main():
	script=app.GetCurrentScript()
	Line = app.GetCalcRelevantObjects('MTM-PIE-2.ElmLne')[0]
	fault_list=['spgf','2pgf']
	#percentages = [10, 20, 40, 70, 85, 77, 98, 100]
	percentages = read_percentages('percentages.csv')[0]
	resistances = read_percentages('percentages.csv')[1]
	currents = {'spgf': [],'2pgf': []}
	xr_ratios = {'spgf': [], '2pgf': []}
	#Set up short circuit 
	shc = app.GetFromStudyCase("shc.ComShc")
	shc.iopt_mde=1 #set mode to IE60909
	shc.Rf = 0 #set fault impedance
	shc.Xf = 0
	shc.shcobj = Line
	for p, r in zip(percentages, resistances):
		for f in fault_list:
			#Set short circuit options - percentage along line and fault type
			shc.ppro = p #set percentage along line
			shc.iopt_shc = f #set fault type
			shc.Rf = r #set fault resistance
			ierr = shc.Execute() #execute short circuit
			xr = get_XR_ratio(Line)
			current = get_3I0(Line)
			currents[f].append(current)
			xr_ratios[f].append(xr)

	#Write results to a csv file
	with open('results.csv', 'w', newline='') as myfile:
		writer = csv.writer(myfile)
		writer.writerow(['Percentage', 'Single Phase to Ground 3I0 (A)', '', '', '2 Phase to Ground 3I0 (A)', '', ''])
		writer.writerow(['', 'Magnitude', 'Angle', 'X/R Ratio', 'Magnitude', 'Angle', 'X/R Ratio'])
		for p, sp, sp_xr, pg2, pg2_xr in zip(percentages, currents['spgf'],xr_ratios['spgf'], currents['2pgf'], xr_ratios['2pgf']):
			writer.writerow([p, 1000*sp[0], sp[1], sp_xr, 1000*pg2[0], pg2[1], pg2_xr]) #Multiplying by 1000 to get the current in amps
			
def get_XR_ratio(line):
	'''Get the X/R ratio at the fault. The X/R ratio is calculated by tan(angle of positive sequence voltage minus angle of positive sequence current'''
	U1_mag_1 = line.GetCubicle(0).GetAttribute('m:U1')
	U1_angle_1 = line.GetCubicle(0).GetAttribute('m:phiu1')
	U1_mag_2 = line.GetCubicle(1).GetAttribute('m:U1')
	U1_angle_2 = line.GetCubicle(1).GetAttribute('m:phiu1')
	U1 = complex_sum([[U1_mag_1, U1_angle_1], [U1_mag_2, U1_angle_2]])

	I1_mag_1 = line.GetAttribute('m:I1:bus1')
	I1_angle_1 = line.GetAttribute('m:phii1:bus1')
	I1_mag_2 = line.GetAttribute('m:I1:bus2')
	I1_angle_2 = line.GetAttribute('m:phii1:bus2')
	I1 = complex_sum([[I1_mag_1, I1_angle_1], [I1_mag_2, I1_angle_2]])
	return np.tan(np.deg2rad(U1[1] - I1[1]))
	
	
def get_3I0(line):
	'''Find 3I0 current at the fault location'''
	I1 = line.GetAttribute('m:I0x3:bus1') #3I0 magnitude
	I2 = line.GetAttribute('m:I0x3:bus2') #3I0 magnitude
	ang1 = line.GetAttribute('m:phii0:bus1') #3I0 angle
	ang2 = line.GetAttribute('m:phii0:bus2') #3I0 angle
	current_sum = complex_sum([[I1, ang1], [I2, ang2]])
	return current_sum
	

def cart2pol(x, y): #returns phi in degrees
	'''Convert complex numbers from cartesian to polar coordinates '''
	rho = np.sqrt(x**2 + y**2)
	phi = np.arctan2(y, x)
	phi = np.rad2deg(phi)
	return(rho, phi)

def pol2cart(rho, phi): #input phi in degrees
	'''Convert complex numbers from polar to cartesian coordinates '''
	phi = np.deg2rad(phi) 
	x = rho * np.cos(phi)
	y = rho * np.sin(phi)
	return(x, y)

def complex_sum(cplx): #Input an array of complex numbers in the form [magnitude, angle]
	'''Returns the sum of an array of complex numbers in polar form'''
	real_sum = 0
	imaj_sum = 0
	for c in cplx:
		real_sum += pol2cart(c[0], c[1])[0]
		imaj_sum += pol2cart(c[0], c[1])[1]
	return cart2pol(real_sum, imaj_sum)

def read_percentages(filename):
	'''
	A function that reads a list of percentage values from a csv file. 
	This method is easier than typing the values in an array in code
	'''
	percentages = []
	resistances = []
	with open(filename, 'r') as myfile:
		reader = csv.reader(myfile)
		for value in reader:
			percentages.append(float(value[0]))
			resistances.append(float(value[1]))

	return [percentages, resistances]


if __name__ == '__main__':
	main()


'''
Fault Types:
3 Phase to Ground: '3psc'
Single phase to ground: 'spgf'
Two phase: '2psc'
Two phase to ground: '2pgf'
'''

