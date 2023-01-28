Build Office.js add-ins with Python and xlwings
-----------------------------------------------

For the full docs, see: https://docs.xlwings.org/en/stable/officejs_addins.html

This quickstart shows you how you can call Python both from a button on the task pane and directly from a Ribbon button. xlwings can be used with any web framework and the quickstart repo therefore contains various implementations such as ``app/server_fastapi.py`` or ``app/server_starlette.py``. At the end of this quickstart, you'll have a working environment for local development.

1. **Download quickstart repo**: Use Git to clone the following repository: https://github.com/xlwings/xlwings-officejs-quickstart. If you don't want to use Git, you could also download the repo by clicking on the green ``Code`` button, followed by ``Download ZIP``, then unzipping it locally.
2. **Update manifest**: If you want to build your own add-in based off this quickstart repo, replace ``<Id>xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx</Id>`` in ``manifest-xlwings-officejs-quickstart.xml`` with a unique ID that you can create by visiting https://www.guidgen.com or by running the following command in Python: ``import uuid;print(uuid.uuid4())``.
3. **Create certificates**: Generate development certificates as the development server needs to be accessed via https instead of http (even on localhost). Otherwise, icons and alerts won't work and Excel on the web won't load the manifest at all. `Download mkcert <https://github.com/FiloSottile/mkcert/releases>`_ (pick the correct file according to your platform), rename the file to ``mkcert``, then run the following commands from a Terminal/Command Prompt (make sure you're in the same directory as ``mkcert``):

   .. code-block:: text

     $ ./mkcert -install
     $ ./mkcert localhost 127.0.0.1 ::1

   This will generate two files ``localhost+2.pem`` and ``localhost+2-key.pem``: move them to the root of the ``xlwings-officejs-quickstart`` quickstart repo.

4. **Install Python dependencies**: 
   
   * Local Python installation: create a virtual or Conda environment and install the Python dependencies by running: ``pip install -r requirements.txt``.
   * Docker: skip this step.
5. **xlwings license key**:

   Get a free `trial license key <https://www.xlwings.org/trial>`_ and install it as follows:

   * Local Python installation: ``xlwings license update -k your-license-key``
   * Docker: set the license key as ``XLWINGS_LICENSE_KEY`` environment variable. The easiest way to do this is to run ``cp .env.template .env`` in a Terminal/Command Prompt and fill in the license key in the ``.env`` file.
6. **Start web app**: 

   * Local Python installation: with the previously created virtual/Conda env activated, start the Python development server by running the Python file with the desired implementation. For example, to run the backend with FastAPI, run the following: ``python app/server_fastapi.py``. You could also run the file via the capabilities offered by your editor.
   * Docker: run ``docker compose up`` instead. Note that Docker by default uses the FastAPI implementation, so you'll need to edit ``docker-compose.yaml`` if you want to change that.
   
   If you see the following, the server is up and running:

   .. code-block:: text

      $ python app/server_fastapi.py 
      INFO:     Will watch for changes in these directories: ['/Users/fz/Dev/xlwings-officejs-quickstart']
      INFO:     Uvicorn running on https://127.0.0.1:8000 (Press CTRL+C to quit)
      INFO:     Started reloader process [56708] using WatchFiles
      INFO:     Started server process [56714]
      INFO:     Waiting for application startup.
      INFO:     Application startup complete.


7. **Sideload the add-in**: Manually load ``manifest-xlwings-officejs-quickstart.xml`` in Excel. This is called *sideloading* and the process differs depending on the platform you're using, see `Office.js docs <https://learn.microsoft.com/en-us/office/dev/add-ins/testing/test-debug-office-add-ins#sideload-an-office-add-in-for-testing>`_ for instructions. Once you've sideloaded the manifest, you'll see the ``Quickstart`` tab in the Ribbon.
8. **Time to play**: You're now ready to play around with the add-in in Excel and make changes to the source code under ``app/server_fastapi.py`` or under the respective file of your framework. Every time you edit and save the Python code, the development server will restart automatically so that you can instantly try out the code changes in Excel. If you make changes to the HTML file, you'll need to right-click on the task pane and select ``Reload``.
