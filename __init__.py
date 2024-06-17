import re
import csv
import logging
import serial
import gtk
import trollius as asyncio
from logging_helpers import _L
from flatland import Integer, Float, Form, Enum, Boolean
from flatland.validation import ValueAtLeast, ValueAtMost
from pygtkhelpers.ui.extra_dialogs import yesno, FormViewDialog
from microdrop.app_context import (get_app, get_hub_uri, MODE_RUNNING_MASK,
                                   MODE_REAL_TIME_MASK)
from microdrop.interfaces import IPlugin
from microdrop.plugin_helpers import (StepOptionsController, AppDataController,
                                      hub_execute)
from microdrop.plugin_manager import (Plugin, implements, PluginGlobals,
                                      ScheduleRequest, emit_signal,
                                      get_service_instance_by_name)
from pygtkhelpers.gthreads import gtk_threadsafe
from pygtkhelpers.ui.dialogs import input
import os.path
import numpy as np
import pandas as pd
import path_helpers as ph
import contextlib
import json
import datetime as dt
import openpyxl as ox
import openpyxl_helpers as oxh

import pelletier_board as pb
import pelletier_board.monitor_dialog as monitor_dialog

from ._version import get_versions
__version__ = get_versions()['version']
del get_versions


logger = logging.getLogger(__name__)


# Add plugin to `"microdrop.managed"` plugin namespace.
PluginGlobals.push_env('microdrop.managed')


class PelletierHeaterPlugin(AppDataController, StepOptionsController,
                                 Plugin):
    '''
    This class is automatically registered with the PluginManager.
    '''
    implements(IPlugin)

    plugin_name = "pelletier_heater_plugin"
    try:
        version = __version__
    except NameError:
        version = 'v0.0.0+unknown'

    AppFields = None

    StepFields = Form.of(Boolean.named('Pelletier_Heater')
                         .using(default=False, optional=True),
                         Float.named('Pelletier_temperature')
                         .using(default=40, optional=True,
                                validators=[ValueAtLeast(minimum=5),
                                            ValueAtMost(maximum=100)]))

    def __init__(self):
          super(PelletierHeaterPlugin, self).__init__()
          # XXX `name` attribute is required in addition to `plugin_name`
          #
          # The `name` attribute is required in addition to the `plugin_name`
          # attribute because MicroDrop uses it for plugin labels in, for
          # example, the plugin manager dialog.
          self.name = self.plugin_name
          self.board = None
          self.monitor = None
          self.heater_is_on = False
          # heater report dialog
          self.hrd = None

    @gtk_threadsafe
    def on_plugin_enable(self):
        '''
        Handler called when plugin is enabled.

        For example, when the MicroDrop application is **launched**, or when
        the plugin is **enabled** from the plugin manager dialog.
        '''
        #Initiate instance of ESP32 class
        self.board = pb.ESP32()
        #get list of ports
        ports = serial.tools.list_ports.comports()

        #try to connect to each port individually, will need a way to test if it's the correct port at some point

        try:
            port = "COM5"
            self.board.connect(port)
            logger.info("Connected to pelletier controller on port: %s", port)

        except Exception as e:
            logger.info("Failed to connect to pelletier controller on port: %s", port)
            self.board = None

        try:
            self.board.send_command('temp')
            self.board.get_data()
            self.heater_report_dialog()
        except:
            pass

        #Warn user if no connection is made
        if not self.board:
            logger.warning("Unable to connect to pelletier controller")
        try:
            super(PelletierHeaterPlugin, self).on_plugin_enable()
        except AttributeError:
            pass

    def on_plugin_disable(self):
        '''
        Handler called when plugin is disabled.

        For example, when the MicroDrop application is **closed**, or when the
        plugin is **disabled** from the plugin manager dialog.
        '''

        try:
            if self.board is not None:
                # Destroy heater report dialog
                if self.hrd is not None:
                    self.hrd.destroy()
                    self.hrd = None
            self.board.close()
        except:
            pass

        try:
            super(PelletierHeaterPlugin, self).on_plugin_disable()
        except AttributeError:
            pass

    @asyncio.coroutine
    def on_step_run(self, plugin_kwargs, signals):
        '''
        Handler called whenever a step is executed  .

        Plugins that handle this signal **MUST** emit the ``on_step_complete``
        signal once they have completed the step.  The protocol controller will
        wait until all plugins have completed the current step before
        proceeding.
        '''
        # Get latest step field values for this plugin.
        self.active_step_kwargs = plugin_kwargs
        options = plugin_kwargs[self.name]

        # Apply step options
        self.apply_step_options(options)

        self.active_step_kwargs = None
        raise asyncio.Return()

    @gtk_threadsafe
    def heater_report_dialog(self):
        if self.hrd is None:
            self.hrd = monitor_dialog._dialog(self.board)


    def start_heater(self, set_point):
        if set_point > 15 :
            self.board.send_command('heater:{}'.format(set_point))
            self.board.send_command('heater_on')
            self.heater_is_on = True
        else:
            self.board.send_command('heater_off')
            self.heater_is_on = False

    def apply_step_options(self, step_options):
        '''
        Apply the specified step options.

        Parameters
        ----------
        step_options : dict
            Dictionary containing the pelletier board plugin options
            for a protocol step.
        '''
        app = get_app()
        app_values = self.get_app_values()

        if self.board:
            step_log = {}

            services_by_name = {service_i.name: service_i
                                for service_i in
                                PluginGlobals
                                .env('microdrop.managed').services}

            step_label = None
            if 'step_label_plugin' in services_by_name:
                # Step label is set for current step
                step_label_plugin = (services_by_name
                                     .get('step_label_plugin'))
                step_label = (step_label_plugin.get_step_options()
                              or {}).get('label')

            # Apply board hardware options.
            try:
                # Heater
                # -------
                if step_options.get('Pelletier_Heater'):
                    heater_target = step_options.get('Pelletier_temperature')

                    self.start_heater(heater_target)

            except Exception:
                logger.error('[%s] Error applying step options.', __name__,
                             exc_info=True)


                app.experiment_log.add_data(step_log, self.name)

        elif not self._user_warned:
            logger.warning('[%s] Cannot apply board settings since board is '
                           'not connected.', __name__, exc_info=True)
            # Do not warn user again until after the next connection attempt.
            self._user_warned = True

    def json_to_excel(self):
    #     '''
    #     Make an excel file from the JSON file
    #     where the temperatures are stored
        app = get_app()
        log_dir = app.experiment_log.get_log_path()

        data_path = log_dir.joinpath('Temperature_log.ndjson')
        if os.path.exists(data_path) == False:
            logger.info('No data to export')
            return

        data = {}
        with open(data_path) as json_file:
            for line in json_file:
                data_json_ij = json.loads(line)
                for key in data_json_ij.keys():
                    try:
                        if key == 'data':
                            data[key].append(data_json_ij[key][0])
                        elif key == 'index':
                            data[key].append(data_json_ij[key])
                    except KeyError:
                        data[key] = []
                        if key == 'data':
                            data[key].append(data_json_ij[key][0])
                        elif key == 'index':
                            data[key].append(data_json_ij[key])

        data['columns'] = data_json_ij['columns']
        data['index'] = np.array(data['index']).flatten()

        df = pd.DataFrame(data['data'],
                  columns = data['columns'],
                  index = data['index'])
        df.index.rename('Time', inplace=True)
        df['DeltaTime(s)'] = (df.index - df.index[0])/1000.0
        df.index = pd.DatetimeIndex(df.index*1e6)
        df = df.reset_index().set_index(['Time','DeltaTime(s)'])

        # output_path = log_dir.joinpath('Temperature_data.xlsx')
        self.dtnow = dt.datetime.now().strftime("%d_%m_%y_%H_%M")
        output_path = log_dir.joinpath('Experiment_log_{}.xlsx'.format(self.dtnow))

        # Create pandas Excel writer to make it easier to append data
        # frames as worksheets.
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            # Configure pandas Excel writer to append to template workbook
            # contents.

            df.to_excel(writer, sheet_name='Temperature_log')

            writer.save()
            writer.close()

    def on_protocol_finished(self):
        # Protocol has finished.  Update
        app_values = self.get_app_values()
        self.json_to_excel()



PluginGlobals.pop_env()
