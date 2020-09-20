import asyncio
import simpleobsws  # type: ignore
import soundcard  # type: ignore

config = {
    'inputs': [
        {
            'obs_device': 'mic-1',
            'device_name': 'Headset Microphone (Rift S)'
        },
        {
            'obs_device': 'mic-2',
            'device_name': 'default'
        }
    ],
    'outputs': [
        {
            'obs_device': 'desktop-1',
            'device_name': 'Speakers (USB Audio Device)'
        }
    ],
    'defaultInput': 'Headset Microphone (Rift S)',
    'defaultOutput': 'Speakers (USB Audio Device)'
}


class AudioDeviceManager(object):
    def __init__(self, config):
        self.default_output = config.get(
            'defaultOutput', soundcard.default_speaker().name)
        self.default_input = config.get(
            'defaultInput', soundcard.default_microphone().name)

    def print_defaults(self):
        print('Default Audio  Input (for this script) set to \'{}\''
              .format(self.default_input))
        print('Default Audio Output (for this script) set to \'{}\''
              .format(self.default_output))

    def select_device(self, prop: str, value: str, is_input=None):
        if is_input is None:
            devices = soundcard.all_speakers() + soundcard.all_microphones()
            if value == 'default':
                return None
        elif is_input:
            if value == 'default':
                return self.select_device('name', self.default_input, True)
            devices = soundcard.all_microphones()
        else:
            if value == 'default':
                return self.select_device('name', self.default_output, True)
            devices = soundcard.all_speakers()
        device = list(filter(lambda d: getattr(d, prop) == value, devices))
        return device[0]

    def get_default(self, is_input: bool):
        return self.select_device('name', 'default', is_input)

    def input_by_name(self, name: str, replace_default=False):
        return self.device_by_name(
            name, is_input=True, replace_default=replace_default)

    def output_by_name(self, name: str, replace_default=False):
        return self.device_by_name(
            name, is_input=False, replace_default=replace_default)

    def device_by_name(self, name: str, is_input=True, replace_default=False):
        if replace_default and name == 'default':
            return self.get_default(is_input)

        try:
            device = self.select_device('name', name, is_input)
        except IndexError:
            print('Windows does not have a device called {}'.format(name))
            self.print_device_listing()
            raise ValueError
        return device

    def device_by_id(self, id: str, default_if_missing=False):
        try:
            return self.select_device('id', id)
        except IndexError:
            if id == 'default':
                return 'default'
            return None

    @classmethod
    def all_devices(self):
        return soundcard.all_speakers() + soundcard.all_microphones()

    @classmethod
    def print_device_listing(self, with_ids=False):
        template = '{1}'
        if with_ids:
            template = 'ID: \'{0}\' Name: \'{1}\''

        speakers = soundcard.all_speakers()
        if len(speakers):
            print('\nAvailable audio output devices:')
            for device in soundcard.all_speakers():
                print(template.format(device.id, device.name))

        mics = soundcard.all_microphones()
        if len(mics):
            print('\nAvailable audio input devices:')
            for device in mics:
                print(template.format(device.id, device.name))


class OBSWebsocket(object):
    def __init__(self, host='172.21.32.1', port=4444):
        loop = asyncio.get_event_loop()
        self.ws = simpleobsws.obsws(host, port, loop=loop)
        self.host = host
        self.port = port
        self.connected = False

    async def connect(self):
        if not self.connected:
            await self.ws.connect()
            self.connected = True
            print('OBS WebSocket connection established.')

    async def disconnect(self):
        if self.connected:
            await self.ws.disconnect()
            self.connected = False
            print('OBS WebSocket disconnected.')

    async def print_sources_settings(self,
                                     audio_device_manager: AudioDeviceManager):
        await self.connect()

        device_names = await self.get_audio_devices()
        for obs_device_name in device_names:
            settings = await self.get_audiodevice_settings(obs_device_name)
            device_id = settings['sourceSettings']['device_id']
            device = audio_device_manager.device_by_id(device_id)
            if not device:
                print('{} is set to a device-id unknown to Windows: \'{}\''
                      .format(obs_device_name, device_id))
                continue
            elif device == 'default':
                print('{} is set to \'default\''.format(obs_device_name))
                continue
            else:
                print('{} is set to \'{}\''
                      .format(obs_device_name, device.name))

    async def get_audio_devices(self):
        await self.connect()
        result = await self.ws.call('GetSpecialSources')
        del result['status']
        return result.values()

    async def get_audio_device_ids(self):
        device_names = await self.get_audio_devices()
        for device_name in device_names:
            settings = await self.get_audiodevice_settings(device_name)
            device_id = settings['sourceSettings']['device_id']
            print('{} is set to \'{}\''.format(device_name, device_id))

    async def get_obs_audio_device_name(self, device):
        await self.connect()
        result = await self.ws.call('GetSpecialSources')
        return result[device]

    async def get_audiodevice_settings(self, device_name='Desktop Audio'):
        await self.connect()
        result = await self.ws.call('GetSourceSettings', {
            'sourceName': device_name,
        })
        return result

    async def call(self, *args, **kwargs):
        await self.connect()
        return await self.ws.call(*args, **kwargs)


async def set_obs_source(ws: OBSWebsocket, source_name: str, device):
    print('Setting OBS source {} to Windows device {}'
          .format(source_name, device.name))
    await ws.call('SetSourceSettings', {
        'sourceName': source_name,
        'sourceSettings': {
            'device_id': device.id
        }
    })


async def set_audio_device_settings(
    ws: OBSWebsocket,
    config: dict,
    adm: AudioDeviceManager
):
    inputs = config['inputs']
    if len(inputs):
        print('\nSetting Input devices')
        for conf in inputs:
            source_name = await ws.get_obs_audio_device_name(
                conf['obs_device'])
            device = adm.input_by_name(conf['device_name'],
                                       replace_default=True)
            await set_obs_source(ws, source_name, device)

    outputs = config['outputs']
    if len(outputs):
        print('\nSetting Output devices')
        for conf in outputs:
            source_name = await ws.get_obs_audio_device_name(
                conf['obs_device'])
            device = adm.output_by_name(conf['device_name'],
                                        replace_default=True)
            await set_obs_source(ws, source_name, device)


async def main():
    AudioDeviceManager.print_device_listing()
    print('')
    adm = AudioDeviceManager(config)
    adm.print_defaults()
    print('')
    ws = OBSWebsocket()
    await ws.connect()
    print('')
    await ws.print_sources_settings(adm)
    await set_audio_device_settings(ws, config, adm)
    print('')
    await ws.disconnect()


if __name__ == "__main__":
    asyncio.run(main())
