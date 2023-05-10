#Verifica stato spool - Ver. 1.0 08-04-19
import time
import win32print

def print_job_checker():
    """
    Stampa documenti in coda ed attende 5 secondi
    """
    jobs = [1]
    while jobs:
        jobs = []
        name = win32print.GetDefaultPrinter()
        phandle = win32print.OpenPrinter(name)
        print_jobs = win32print.EnumJobs(phandle, 0, -1, 1)
        if print_jobs:
            jobs.extend(list(print_jobs))
        for job in print_jobs:
            document = job["pDocument"]
            print "Documento in coda => " + document
        win32print.ClosePrinter(phandle)
        time.sleep(5)
    print "Non ci sono documenti in coda!"

if __name__ == "__main__":
    print_job_checker()
