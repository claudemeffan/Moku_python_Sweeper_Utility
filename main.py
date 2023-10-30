from PyQt6 import QtWidgets
import pyqtgraph as pg
from moku.instruments import LockInAmp
from moku.exceptions import InvalidParameterException
from PyQt6 import QtGui, QtCore
import numpy as np
import subprocess
from datetime import datetime
import pandas as pd


class FrequencySweepGUI(QtWidgets.QApplication):
    def __init__(self):
        super(FrequencySweepGUI, self).__init__([])

        # LIA object
        self.lia = None

        # Current data object (empty)
        self.cur = ([0], [0])
        self.cur_name = None

        # Basic Window Set up
        self.w = QtWidgets.QWidget()
        self.w.setWindowTitle('Moku Frequency Sweeper')

        ## Create a grid layout to manage the widgets size and position
        self.layout = QtWidgets.QGridLayout()
        self.w.setLayout(self.layout)

        ## Connect to Moku Button + config Positions 0-1
        self.connect_btn = QtWidgets.QPushButton('Connect')
        self.connect_btn.clicked.connect(self.connect_and_setup)
        self.connection_status = QtWidgets.QLabel('Status: No Connection...')
        self.available_inst = QtWidgets.QComboBox()
        self.available_inst.setMaximumWidth(150)
        self.populate_available_instruments()

        ## Add widgets to the layout
        self.layout.addWidget(QtWidgets.QLabel('Available IPs:'), 0, 0)
        self.layout.addWidget(self.available_inst, 1, 0)
        self.layout.addWidget(self.connect_btn, 1, 1)  # button goes in upper-left
        self.layout.addWidget(self.connection_status, 0, 1)

        # Input Settings
        '''
        self.layout.addWidget(QtWidgets.QLabel('Input Attenuation:'), 2, 0)
        self.layout.addWidget(QtWidgets.QLabel('Input Coupling:'), 3, 0)
        self.layout.addWidget(QtWidgets.QLabel('Input Impedance'), 4, 0)

        self.input_attenuation = QtWidgets.QComboBox()
        self.input_attenuation.addItems(['0dB', '-14dB'])
        self.input_coupling = QtWidgets.QComboBox()
        self.input_coupling.addItems(['AC', 'DC'])
        self.input_impedance = QtWidgets.QComboBox()
        self.input_impedance.addItems(['50Ohm', '1MOhm'])

        self.layout.addWidget(self.input_attenuation, 2, 1)
        self.layout.addWidget(self.input_coupling, 3, 1)
        self.layout.addWidget(self.input_impedance)
        '''

        ## Frequency Sweep Parameters 2-4
        self.layout.addWidget(QtWidgets.QLabel('Freq. Start [Hz]:'), 5, 0)
        self.layout.addWidget(QtWidgets.QLabel('Freq. Stop [Hz]:'), 6, 0)
        self.layout.addWidget(QtWidgets.QLabel('Points'), 7, 0)

        self.f_start = QtWidgets.QDoubleSpinBox(maximum=int(600e6), value=1000)
        self.f_stop = QtWidgets.QDoubleSpinBox(maximum=int(600e6), value=10000)
        self.f_point = QtWidgets.QSpinBox(maximum=int(4000), value=50)  # Maximum points per sweep is 4000

        self.layout.addWidget(self.f_start, 5, 1)
        self.layout.addWidget(self.f_stop, 6, 1)
        self.layout.addWidget(self.f_point, 7, 1)

        ## Set up Plotting Outputs 5-8
        self.output_select = QtWidgets.QComboBox()
        self.output_select.addItems(['R', 'Theta', 'X', 'Y'])
        self.layout.addWidget(QtWidgets.QLabel('Measurement Select:'), 8, 0)
        self.layout.addWidget(self.output_select, 8, 1)

        ## Set up Plotting Outputs 5-8
        self.output_voltage = QtWidgets.QDoubleSpinBox(maximum=10.0, value=0.5)
        self.layout.addWidget(QtWidgets.QLabel('Output 2 Drive [Vpp]:'), 9, 0)
        self.layout.addWidget(self.output_voltage, 9, 1)

        self.run_sweep_button = QtWidgets.QPushButton('Run Sweep')
        self.layout.addWidget(self.run_sweep_button, 10, 1)
        self.run_sweep_button.clicked.connect(self.run_sweep)

        ## Set up data sweep tree
        self.data_plotted = pg.DataTreeWidget()
        self.data_dict = {}
        self.layout.addWidget(self.data_plotted, 11, 0, 5, 2)
        self.data_plotted.setHeaderLabel('Displayed Data')

        # Delete button for tree items
        self.delete_button = QtWidgets.QPushButton('Delete')
        self.layout.addWidget(self.delete_button, 17, 0)
        self.delete_button.clicked.connect(self.deleteEntry)

        ## Save Data Button setup
        self.save_data_button = QtWidgets.QPushButton('Save data')
        self.layout.addWidget(self.save_data_button, 17, 1)
        self.save_data_button.clicked.connect(self.save_data)

        # Plot Item Set-up
        self.plot = pg.PlotWidget()
        self.layout.addWidget(self.plot, 0, 2, 20, 2)  # plot goes on right side, spanning 3 rows
        # Display the widget as a new window
        self.setup_plot()

        ## Show the window
        self.w.show()

        ## Start the Qt event loop
        self.exec()


    def get_settings(self):
        self.lia.get
        pass

    def deleteEntry(self):
        x = self.data_plotted.selectedItems()
        if not x:
            return
        name = x[0].text(0)
        del self.data_dict[name]
        # Clear last recorded trace if nessecary.
        if name == self.cur_name:
            self.cur = ([0], [0])
            self.cur_name = ''
        self.updateTree()
        self.refresh_plot()

    def populate_available_instruments(self):
        result = subprocess.run(['moku', 'list'], stdout=subprocess.PIPE)
        found_devices = result.stdout.decode('utf-8').split('\n')[3:-1]
        self.available_inst.addItems(found_devices)

    def save_data(self):
        if len(self.data_dict) == 0:
            return
        save_file = QtWidgets.QFileDialog.getSaveFileName(filter="Excel files (*.xlsx)")
        df = pd.DataFrame()
        for measurement in self.data_dict:
            f = self.data_dict[measurement][0]
            v = self.data_dict[measurement][1]
            df[f'{measurement} f'] = f
            df[f'{measurement}'] = v
        df.to_excel(save_file[0])

    def connect_and_setup(self):
        try:
            ip = self.available_inst.currentText().split('    ')[-1]
            self.setOverrideCursor(QtGui.QCursor(QtCore.Qt.CursorShape.WaitCursor))
            self.lia = LockInAmp(f'[{ip}]', force_connect=True)
            self.lia.set_timebase(0, 100e-6, strict=False)  # Can make averaging times adjustable in the future
            self.restoreOverrideCursor()
            self.connection_status.setText('Status: Connected!')
            self.connection_status.update()

        except:
            self.restoreOverrideCursor()
            message_box = QtWidgets.QMessageBox()
            message_box.setIcon(QtWidgets.QMessageBox.Icon.Warning)  # This doesn't work for some reason.
            message_box.setWindowTitle("Connection Failed!")
            message_box.setText("Could not connect to Moku. Please check IP address and try again")
            # Add a "Clear" button
            message_box.exec()

    def setup_plot(self):
        # Set the label for the x-axis
        self.plot.getAxis("bottom").setLabel("Frequency [Hz]")
        # Set the label for the y-axis
        self.plot.getAxis("left").setLabel("Amplitude [Volts]")

    def run_sweep(self):
        if self.lia is None:
            message_box = QtWidgets.QMessageBox()
            message_box.setIcon(QtWidgets.QMessageBox.Icon.Warning)  # This doesn't work for some reason.
            message_box.setWindowTitle("No Moku Connected")
            message_box.setText("Not Connected to any Moku Device. Please try again.")
            # Add a "Clear" button
            message_box.exec()

        frequencies = np.linspace(self.f_start.value(),
                                  self.f_stop.value(),
                                  self.f_point.value())

        self.lia.set_monitor(1, 'MainOutput')  # Sets the main output
        self.lia.set_outputs(main=f"{self.output_select.currentText()}",
                             main_offset=0,
                             aux="Demod",
                             aux_offset=0,
                             strict=False)


        # Convert voltage into a specifc gain
        self.lia.set_aux_output(frequency=1000, amplitude=self.output_voltage.value()) # Frequency Value

        new_data = []  # Empty list to add new points into

        # LPF Settings
        self.lia.set_filter(corner_frequency=int(100), slope="Slope24dB", strict=False)

        loading_dialog = pg.ProgressDialog('Running sweep...',
                                           minimum=0,
                                           maximum=self.f_point.value(),
                                           cancelText='Cancel',
                                           wait=0)

        i = 0  # Counter for the loading dialog
        for f in frequencies:
            try:
                self.lia.set_demodulation(mode="Internal", frequency=f)

            except InvalidParameterException as error:
                message_box = QtWidgets.QMessageBox()
                message_box.setIcon(QtWidgets.QMessageBox.Icon.Warning)  # This doesn't work for some reason.
                message_box.setWindowTitle("Parameter Setting Failed")
                message_box.setText(error.args[0][0])
                # Add a "Clear" button
                message_box.exec()
                break

            if loading_dialog.wasCanceled():
                break

            i += 1
            loading_dialog.setValue(i)

            data = np.mean(self.lia.get_data()['ch1']) #wait_reacquire=True Maybe need to add this later..
            new_data.append(data)

        # update the tree data structure to contain the new data.
        self.cur = (frequencies, np.array(new_data))
        self.cur_name = f'{self.output_select.currentText()}    {datetime.now()}'
        self.data_dict[self.cur_name] = self.cur
        self.updateTree()
        self.refresh_plot()

    def updateTree(self):
        self.data_plotted.setData(self.data_dict, hideRoot=True)
        self.collapseAllItems(self.data_plotted.invisibleRootItem())

    def collapseAllItems(self, item):
        item.setExpanded(False)  # Collapse the current item
        for i in range(item.childCount()):
            child = item.child(i)
            self.collapseAllItems(child)

    def refresh_plot(self):
        # Clears the plot
        self.plot.clear()
        #Plots all old data
        for measurement in self.data_dict:
            self.plot.plot(x=self.data_dict[measurement][0], y=self.data_dict[measurement][1])
        # Separate plot command with color for most recent data
        self.plot.plot(x=self.cur[0], y=self.cur[1], pen='r') # Replace this with the most rec
        # Last measurement


FrequencySweepGUI()
