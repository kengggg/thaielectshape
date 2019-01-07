# this version has design for google cloud function

import shapefile
from pyproj import Proj, transform
import pandas as pd
from collections import OrderedDict
from json import dumps
from google.cloud import storage
from google.cloud import bigquery
import fiona
import zipfile
import os
import glob
import sys
import re

def csvtobq(source_file, **kwargs):
    replace = kwargs.get('replace', False)
    table_id = kwargs.get('table_id')
    dataset_id = 'demos'
    bigquery_client = bigquery.Client()
    dataset_ref = bigquery_client.dataset(dataset_id)
    # This example uses JSON, but you can use other formats.
    # See https://cloud.google.com/bigquery/loading-data
    job_config = bigquery.LoadJobConfig()
    job_config.source_format = bigquery.SourceFormat.CSV
    job_config.write_disposition = bigquery.WriteDisposition.WRITE_TRUNCATE
    job_config.autodetect = True

    job = bigquery_client.load_table_from_file(
        source_file, dataset_ref.table(table_id), job_config=job_config)
    job.result()
    print(f'Loaded {job.output_rows} rows into {dataset_id}:{table_id} {job.state}.')


def loadtobq(csvfilename, shpname):
    with open(csvfilename, mode='rb') as csvfile:
        csvtobq(csvfile, table_id=shpname)
    print(f'LOAD to demos.{shpname} SUCCESS')

def reproject(geom,crs='EPSG:32647'):
    # Define dictionary representation of output feature collection
    from_crs = Proj(init=crs)
    to_crs = Proj(init='epsg:4326')
    # Iterate through each feature of the feature collection
    new_coords = []
    for i,shape in enumerate(geom['coordinates']):
        x, y = transform(from_crs, to_crs, *zip(*shape))
        new_coords.append([list(a) for a in zip(x, y)])
    geom['coordinates']=new_coords
    return geom

def shptodf(shppath):
    shpfilename = os.path.join(os.path.dirname(shppath),
                               glob.glob1(os.path.dirname(shppath), '*.shp')[0])
    shpname = os.path.basename(os.path.splitext(shpfilename)[0])
    with fiona.open(shppath, 'r') as source:
        crs=source.crs
    try:
        reader = shapefile.Reader(shpfilename, encoding='utf-8')
        shapeRecords = reader.shapeRecords()
    except:
        reader = shapefile.Reader(shpfilename, encoding="ISO-8859-11")
        shapeRecords = reader.shapeRecords()
    fields = reader.fields[1:]
    field_names = [field[0] for field in fields]
    buffer = []
    for sr in shapeRecords:
        atr = dict(zip(field_names, sr.record))
        geom = sr.shape.__geo_interface__
        if crs['init'] != 'epsg:4326':
            geom = reproject(geom,crs['init'])
        row = OrderedDict()
        row["geom"] = dumps(geom)
        row.update(atr)
        buffer.append(row)
    reader.close()
    return pd.DataFrame(buffer)

def shptocsv(shppath):
    df=shptodf(shppath)
    shpfilepath = glob.glob1(os.path.dirname(shppath), '*.shp')[0]
    csvfilename=shpfilepath[:-3]+'csv'
    df.to_csv(csvfilename, index=False)
    return csvfilename


def main(data, context):
    if data['name'].endswith('.zip'):
        storage_client=storage.Client()
        bucket=storage_client.bucket(data['bucket'])
        blob=bucket.blob(data['name'])
        zipfilepath=os.path.join('/tmp', data['name'])
        blob.download_to_filename(zipfilepath)
        #unzip file
        with zipfile.ZipFile(zipfilepath, 'r') as zip_ref:
            test_result=zip_ref.testzip()
            if test_result:
                print(f"First bad file in zip: ")
                sys.exit(1)
            rootdirzips=[x for x in zip_ref.namelist() if re.match('^[\w\s]+\/$', x) and
                         x not in ['__MACOSX/', '.DS_Store/']]
            glob.glob1(zipfilepath, '*')
            zip_ref.extractall(os.path.dirname(zipfilepath))

        if rootdirzips:
            for rootdirzip in rootdirzips:
                shppath=os.path.join(zipfilepath,rootdirzip)
                shptocsv(shppath)
        else:
            shppath = zipfilepath
            shptocsv(shppath)