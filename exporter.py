import clr
import time
import os
import logging
import sys
import socket

import win32serviceutil

import servicemanager
import win32event
import win32service

from prometheus_client import start_http_server, Gauge

log = logging.getLogger(__name__)
log.setLevel(logging.INFO)
log.addHandler(logging.StreamHandler(sys.stdout))

PORT=8000
SCRAPE_INTERVAL=1
DLL_PATH='C:\\Program Files\\HardwareInfoExporter\\OpenHardwareMonitorLib.dll'

class HostMetricsWrapper(object):
    def __init__(self):
        self._handler = None
        self.sensor_type = ['Voltage', 'Clock', 'Temperature', 'Load', 'Fan', 'Flow', 'Control', 'Level', 'Factor', 'Power', 'Data', 'SmallData', 'Throughput']
        self.openhardwaremonitor_hwtypes = ['Mainboard','SuperIO','CPU','RAM','GpuNvidia','GpuAti','TBalancer','Heatmaster','HDD']

        self._metrics = {
            'cpu': {},
            'gpu': {},
            'disk': {},
            'memory': {},
        }

    def init_open_hw_monitor(self):
        clr.AddReference(DLL_PATH)

        from OpenHardwareMonitor import Hardware

        handle = Hardware.Computer()
        handle.MainboardEnabled = True
        handle.CPUEnabled = True
        handle.RAMEnabled = True
        handle.GPUEnabled = True
        handle.HDDEnabled = True
        handle.Open()
        self._handler = handle
    
    def get_all_metrics(self):
        for i in self._handler.Hardware:
            i.Update()
            for sensor in i.Sensors:
                self.parse_sensors(sensor)
                for j in i.SubHardware:
                    for sensor in j.Sensors:
                        self.parse_sensors(sensor)
        return self._metrics

    def parse_sensors(self, sensor):
        hw_type = self.openhardwaremonitor_hwtypes[sensor.Hardware.HardwareType]
        type_to_func_mapping = {
            'CPU': self.load_cpu_metrics,
            'RAM': self.load_memory_metrics,
            'GpuNvidia': self.load_gpu_metrics,
            'HDD': self.load_disk_metrics
        }
        type_to_func_mapping[hw_type](sensor)
        if hw_type not in ['CPU', 'RAM', 'GpuNvidia', 'HDD']:
            log.warn(f"This hardware type has no data collected [{hw_type}]")

    def load_cpu_metrics(self, sensor):
        core_number = sensor.Name[sensor.Name.rfind('#')+1:]
        stype = self.sensor_type[sensor.SensorType]
        name = sensor.Hardware.Name

        if self._metrics['cpu'].get(name) is None: self._metrics['cpu'][name] = {}
        if self._metrics['cpu'][name].get(stype) is None: self._metrics['cpu'][name][stype] = {}

        if stype == "Power":
            if sensor.Name == 'CPU Package': self._metrics['cpu'][name][stype]['Package'] = sensor.Value
            if sensor.Name == 'CPU Cores': self._metrics['cpu'][name][stype]['Cores'] = sensor.Value
            if sensor.Name == 'CPU Graphics': self._metrics['cpu'][name][stype]['Graphics'] = sensor.Value
            if sensor.Name == 'CPU DRAM': self._metrics['cpu'][name][stype]['DRAM'] = sensor.Value
        else:
            self._metrics['cpu'][name][stype][core_number] = sensor.Value

    def load_cpu_power_metrics(self, sensor):
        stype = self.sensor_type[sensor.SensorType]
        name = sensor.Hardware.Name

        if self._metrics['cpu'].get(name) is None: self._metrics['cpu'][name] = {}
        if self._metrics['cpu'][name].get(stype) is None: self._metrics['cpu'][name][stype] = {}

        if sensor.Name == 'CPU Package': self._metrics['cpu'][name][stype]['Package'] = sensor.Value
        if sensor.Name == 'CPU Cores': self._metrics['cpu'][name][stype]['Cores'] = sensor.Value
        if sensor.Name == 'CPU Graphics': self._metrics['cpu'][name][stype]['Graphics'] = sensor.Value
        if sensor.Name == 'CPU DRAM': self._metrics['cpu'][name][stype]['DRAM'] = sensor.Value

    def load_memory_metrics(self, sensor):
        stype = self.sensor_type[sensor.SensorType]
        name = sensor.Hardware.Name

        if self._metrics['memory'].get(name) is None: self._metrics['memory'][name] = {}
        if self._metrics['memory'][name].get(stype) is None: self._metrics['memory'][name][stype] = {}

        if sensor.Name == 'Memory': self._metrics['memory'][name][stype]['usage'] = sensor.Value
        if sensor.Name == 'Used Memory': self._metrics['memory'][name][stype]['used'] = sensor.Value
        if sensor.Name == 'Available Memory': self._metrics['memory'][name][stype]['free'] = sensor.Value

    def load_gpu_metrics(self, sensor):
        stype = self.sensor_type[sensor.SensorType]
        name = sensor.Hardware.Name

        if self._metrics['gpu'].get(name) is None: self._metrics['gpu'][name] = {}
        if self._metrics['gpu'][name].get(stype) is None: self._metrics['gpu'][name][stype] = {}
        if self._metrics['gpu'][name].get('Memory') is None: self._metrics['gpu'][name]['Memory'] = {}

        if sensor.Name == 'GPU Core': self._metrics['gpu'][name][stype]['GPU Core'] = sensor.Value
        if sensor.Name == 'GPU Memory Total': self._metrics['gpu'][name]['Memory']['total'] = sensor.Value
        if sensor.Name == 'GPU Memory Used': self._metrics['gpu'][name]['Memory']['used'] = sensor.Value
        if sensor.Name == 'GPU Memory Free': self._metrics['gpu'][name]['Memory']['free'] = sensor.Value

        if sensor.Name == 'GPU Frame Buffer': self._metrics['gpu'][name]['Load']['Frame Buffer'] = sensor.Value
        if sensor.Name == 'GPU Video Engine': self._metrics['gpu'][name]['Load']['Video Engine'] = sensor.Value
        if sensor.Name == 'GPU Bus Interface': self._metrics['gpu'][name]['Load']['Bus Interface'] = sensor.Value

        if sensor.Name == 'GPU Memory': self._metrics['gpu'][name]['Clock']['Memory'] = sensor.Value
        if sensor.Name == 'GPU Shader': self._metrics['gpu'][name]['Clock']['Shader'] = sensor.Value
        if sensor.Name == 'GPU': self._metrics['gpu'][name]['Fan']['1'] = sensor.Value

    def load_disk_metrics(self, sensor):
        stype = self.sensor_type[sensor.SensorType]
        name = sensor.Hardware.Name

        if self._metrics['disk'].get(name) is None: self._metrics['disk'][name] = {}
        if self._metrics['disk'][name].get(stype, None) is None: self._metrics['disk'][name][stype] = {}
        self._metrics['disk'][name][stype] = sensor.Value


class HardwareMetricsExporter(object):
    def __init__(self, port):
        start_http_server(port)

        self._cpu_temperature = Gauge('cpu_temperature', 'cpu_temperature', ['device', 'core'])
        self._cpu_load = Gauge('cpu_load', 'cpu_load', ['device', 'core'])
        self._cpu_clock = Gauge('cpu_clock', 'cpu_clock',  ['device', 'core'])
        self._cpu_power = Gauge('cpu_power', 'cpu_power',  ['device', 'core'])

        self._gpu_temperature = Gauge('gpu_temperature', 'cpu_temperature', ['device', 'block'])
        self._gpu_memory = Gauge('gpu_memory', 'gpu_memory',  ['device', 'block'])
        self._gpu_clock = Gauge('gpu_clock', 'cpu_clock',  ['device', 'block'])
        self._gpu_fan = Gauge('gpu_fan', 'gpu_fan',  ['device', 'block'])
        self._gpu_load = Gauge('gpu_load', 'cpu_load', ['device', 'block'])

        self._disk_load = Gauge('disk_load', 'disk_load', ['device'])

        self._memory_load = Gauge('memory_load', 'memory_load', ['device'])
        self._memory_used = Gauge('memory_used', 'memory_used', ['device'])
        self._memory_free = Gauge('memory_free', 'memory_free', ['device'])

    def update(self, all_metrics):
        self._update_cpu_metrics(all_metrics)
        self._update_gpu_metrics(all_metrics)
        self._update_memory_metrics(all_metrics)
        self._update_disk_metrics(all_metrics)

    def _update_cpu_metrics(self, all_metrics):
        for name, v in all_metrics['cpu'].items():
            for metric_type, metrics in v.items():
                cpu_metric_to_func = {
                    'Load': self._cpu_load,
                    'Temperature': self._cpu_temperature,
                    'Clock': self._cpu_clock,
                    'Power': self._cpu_power,
                }
                for spec, value in metrics.items():
                    cpu_metric_to_func[metric_type].labels(name, f"{spec}").set(value)

    def _update_gpu_metrics(self, all_metrics):
        for name, v in all_metrics['gpu'].items():
            for metric_type, metrics in v.items():
                gpu_metric_to_func = {
                    'Load': self._gpu_load,
                    'Temperature': self._gpu_temperature,
                    'Clock': self._gpu_clock,
                    'Fan': self._gpu_fan,
                    'Memory': self._gpu_memory,
                }
                for spec, value in metrics.items():
                    gpu_metric_to_func[metric_type].labels(name, f"{spec}").set(value)

    def _update_memory_metrics(self, all_metrics):
        for name, v in all_metrics['memory'].items():
            for metric_type, metrics in v.items():
                for spec, value in metrics.items():
                    if metric_type == 'Load':
                        self._memory_load.labels(name).set(value)
                    if metric_type == 'Data':
                        mapp = {
                            'used': self._memory_used,
                            'free': self._memory_free
                        }
                        mapp[spec].labels(name).set(value)

    def _update_disk_metrics(self, all_metrics):
        for name, v in all_metrics['disk'].items():
            for _, value in v.items():
                    self._disk_load.labels(name).set(value)


class HardwareInfoExporter(win32serviceutil.ServiceFramework):
    '''Base class to create winservice in Python'''

    _svc_name_ = 'HardwareInfoExporter'
    _svc_display_name_ = 'HardwareInfoExporter'
    _svc_description_ = 'HardwareInfoExporter'

    @classmethod
    def parse_command_line(cls):
        '''
        ClassMethod to parse the command line
        '''
        win32serviceutil.HandleCommandLine(cls)

    def __init__(self, args):
        '''
        Constructor of the winservice
        '''
        self._running = True
        win32serviceutil.ServiceFramework.__init__(self, args)
        self.hWaitStop = win32event.CreateEvent(None, 0, 0, None)
        socket.setdefaulttimeout(60)

    def SvcStop(self):
        '''
        Called when the service is asked to stop
        '''
        self.stop()
        self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)
        win32event.SetEvent(self.hWaitStop)

    def SvcDoRun(self):
        '''
        Called when the service is asked to start
        '''
        self.start()
        servicemanager.LogMsg(servicemanager.EVENTLOG_INFORMATION_TYPE,
                              servicemanager.PYS_SERVICE_STARTED,
                              (self._svc_name_, ''))
        self.main()

    def start(self):
        '''
        Override to add logic before the start
        eg. running condition
        '''
        self._running = True

    def stop(self):
        '''
        Override to add logic before the stop
        eg. invalidating running condition
        '''
        self._running = False

    def main(self):
        host_metrics = HostMetricsWrapper()
        host_metrics.init_open_hw_monitor()

        exporter = HardwareMetricsExporter(PORT)
        while(True):
            metrics = host_metrics.get_all_metrics()
            exporter.update(metrics)
            log.info('Metrics collected.')
            time.sleep(SCRAPE_INTERVAL)
            if not self._running:
                break


if __name__ == '__main__':
    if len(sys.argv) == 1:
        servicemanager.Initialize()
        servicemanager.PrepareToHostSingle(HardwareInfoExporter)
        servicemanager.StartServiceCtrlDispatcher()
    else:
        win32serviceutil.HandleCommandLine(HardwareInfoExporter)
