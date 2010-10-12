Query and update an Inteum MSSQL database via email using Lotus Notes
=====================================================================
Inteum-Lotus-Helper is a mail-based interface to searching and updating Inteum's underlying MSSQL database via Lotus Notes.  Inteum is a software system for managing intellectual property.  Authorized users send commands via email to search the database, retrieve documents, store remarks with attached documents and reserve case numbers.  Inteum-Lotus-Helper runs as a process in the background, listening to its assigned Lotus Notes mailbox for new email messages containing commands.


Install dependencies
--------------------
1. Install `pywin32 <http://sourceforge.net/projects/pywin32>`_

2. Install `pyodbc <http://code.google.com/p/pyodbc>`_


Configure tool
--------------
1. Copy ``default.cfg`` to ``.default.cfg``

2. Complete all fields in the configuration file ``.default.cfg``

3. To suppress "Begin CD to MIME Conversion" console messages, append ``converter_log_level=10`` to all instances of ``notes.ini`` on the machine.


Check that the tool works properly
----------------------------------
1. Send an email to the mailbox with a question mark as the subject.

2. Run ``python go.py - 1``.


Schedule the tool to run regularly
----------------------------------
1. To run Inteum-Lotus-Helper as a background process, use ``All Programs > Accessories > System Tools > Scheduled Tasks`` to schedule ``pythonw go.py`` to run every thirty minutes.  Make sure you specify ``pythonw`` and not ``python`` to prevent the black console from appearing on every run.  Marking "Run only if logged in" will enable the task to run even when you change your Windows password; however, if you choose this option, you must remain logged into your computer for the task to run.

2. To schedule a summary spreadsheet to be sent to your mailbox, use ``All Programs > Accessories > System Tools > Scheduled Tasks`` to schedule ``pythonw report.py`` to run every morning.
