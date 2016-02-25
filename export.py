class tExportCSVPM(QtCore.QThread):
    def __init__(self, parent = None):
        QtCore.QThread.__init__(self, parent)
        self.signals = custSignal()
        
        self.data = []
        self.data.append(["FM", "KD_JENIS_TRANSAKSI", "FG_PENGGANTI", "NOMOR_FAKTUR", "MASA_PAJAK", "TAHUN_PAJAK", "TANGGAL_FAKTUR", "NPWP", "NAMA", "ALAMAT_LENGKAP", "JUMLAH_DPP", "JUMLAH_PPN", "JUMLAH_PPNBM", "IS_CREDITABLE"])

    def run(self):
        self.signals.expStart.emit('Starting...')
        ok = False
        try:
            self.signals.expLoadData.emit('Loading data...')
            if exp_all == 0:
                url = 'http://'+svc_host+':'+svc_port+'/pm_list?where={"myNPWP":"like('+'\\"'+str(exp_npwp)+'\\'+'")","TAHUN_PAJAK":"like('+'\\"'+str(exp_tahun)+'\\'+'")","MASA_PAJAK":"like('+'\\"'+str(exp_masa)+'\\'+'")","exported":"like('+'\\"'+str(exp_all)+'\\'+'")"}'
            else:
                url = 'http://'+svc_host+':'+svc_port+'/pm_list?where={"myNPWP":"like('+'\\"'+str(exp_npwp)+'\\'+'")","TAHUN_PAJAK":"like('+'\\"'+str(exp_tahun)+'\\'+'")","MASA_PAJAK":"like('+'\\"'+str(exp_masa)+'\\'+'")"}'
            r = requests.get(url,auth = (authresponse['token'],''))
            data = json.loads(r.content)
            if '_meta' in data and data['_meta']['total'] > 25:
                if exp_all == 0:
                    url = 'http://'+svc_host+':'+svc_port+'/pm_list?max_results='+str(data['_meta']['total'])+'&where={"myNPWP":"like('+'\\"'+str(exp_npwp)+'\\'+'")","TAHUN_PAJAK":"like('+'\\"'+str(exp_tahun)+'\\'+'")","MASA_PAJAK":"like('+'\\"'+str(exp_masa)+'\\'+'")","exported":"like('+'\\"'+str(exp_all)+'\\'+'")"}'
                else:
                    url = 'http://'+svc_host+':'+svc_port+'/pm_list?max_results='+str(data['_meta']['total'])+'&where={"myNPWP":"like('+'\\"'+str(exp_npwp)+'\\'+'")","TAHUN_PAJAK":"like('+'\\"'+str(exp_tahun)+'\\'+'")","MASA_PAJAK":"like('+'\\"'+str(exp_masa)+'\\'+'")"}'
                r = requests.get(url,auth = (authresponse['token'],''))
                data = json.loads(r.content)
            if '_items' in data:
                r_items = data['_items']
                for i in r_items:
                    r_list = []
                    r_list.append(str(i['FM']))
                    r_list.append(str(i['KD_JENIS_TRANSAKSI']))
                    r_list.append(str(i['FG_PENGGANTI']))
                    r_list.append(str(i['NOMOR_FAKTUR']))
                    r_list.append(str(i['MASA_PAJAK']))
                    r_list.append(str(i['TAHUN_PAJAK']))
                    tanggalFaktur = i['TANGGAL_FAKTUR']
                    _tanggal, _bulan, _tahun = tanggalFaktur.split('/')                
                    if _tanggal[0] == '0':
                        _tanggal = _tanggal[1]
                    if _bulan[0] == '0':
                        _bulan = _bulan[1]
                    _tanggalFaktur = _tanggal + '/' + _bulan + '/' + _tahun
                    r_list.append(str(_tanggalFaktur))
                    r_list.append(str(i['NPWP']))
                    r_list.append(str(i['NAMA']))
                    r_list.append(str(i['ALAMAT_LENGKAP']))
                    r_list.append(str(i['JUMLAH_DPP']))
                    r_list.append(str(i['JUMLAH_PPN']))
                    r_list.append(str(i['JUMLAH_PPNBM']))
                    r_list.append(str(i['IS_CREDITABLE']))
                    self.data.append(r_list)
                
                ok = True
        except:
            pass
            
        if ok:
            self.signals.expWrite.emit('Writing data...')
            with open(exp_filepath, 'wb') as myfile:
                wr = csv.writer(myfile, quoting=csv.QUOTE_ALL)
                row = 0
                for values in self.data:
                    wr.writerow(values)
                    row += 1
                    self.signals.expPercent.emit(float(row)/float(len(self.data))*100)
            try:
                url = 'http://'+svc_host+':'+svc_port+'/pm_list?where={"myNPWP":"like('+'\\"'+str(exp_npwp)+'\\'+'")","TAHUN_PAJAK":"like('+'\\"'+str(exp_tahun)+'\\'+'")","MASA_PAJAK":"like('+'\\"'+str(exp_masa)+'\\'+'")"}'
                r = requests.get(url,auth = (authresponse['token'],''))
                data = json.loads(r.content)
                if '_meta' in data and data['_meta']['total'] > 25:
                    url = 'http://'+svc_host+':'+svc_port+'/pm_list?max_results='+str(data['_meta']['total'])+'&where={"myNPWP":"like('+'\\"'+str(exp_npwp)+'\\'+'")","TAHUN_PAJAK":"like('+'\\"'+str(exp_tahun)+'\\'+'")","MASA_PAJAK":"like('+'\\"'+str(exp_masa)+'\\'+'")"}'
                    r = requests.get(url,auth = (authresponse['token'],''))
                    data = json.loads(r.content)
                if '_items' in data:
                    r_items = data['_items']
                    payload = {'exported':1}
                    for r_item in r_items :
                        link = r_item['_links']['self']['href']
                        etag = r_item['_etag']
                
                        url = 'http://'+svc_host+':'+svc_port+'/'+link
                        headers = {'content-type':'application/json','if-match':etag}
                        r = requests.patch(url,data = json.dumps(payload),headers = headers,auth = (authresponse['token'],''))
            except:
                pass
            self.signals.expFinish.emit(True)
        else:
            self.signals.expLoadData.emit('Failed...')
class tExportExcelPM(QtCore.QThread):
    def __init__(self, parent = None):
        QtCore.QThread.__init__(self, parent)
        self.signals = custSignal()
        
        self.data = []
        self.data.append(["FM", "KD_JENIS_TRANSAKSI", "FG_PENGGANTI", "NOMOR_FAKTUR", "MASA_PAJAK", "TAHUN_PAJAK", "TANGGAL_FAKTUR", "NPWP", "NAMA", "ALAMAT_LENGKAP", "JUMLAH_DPP", "JUMLAH_PPN", "JUMLAH_PPNBM", "IS_CREDITABLE"])

    def run(self):
        self.signals.expStart.emit('Starting...')
        ok = False
        try:
            self.signals.expLoadData.emit('Loading data...')
            if exp_all == 0:
                url = 'http://'+svc_host+':'+svc_port+'/pm_list?where={"myNPWP":"like('+'\\"'+str(exp_npwp)+'\\'+'")","TAHUN_PAJAK":"like('+'\\"'+str(exp_tahun)+'\\'+'")","MASA_PAJAK":"like('+'\\"'+str(exp_masa)+'\\'+'")","exported":"like('+'\\"'+str(exp_all)+'\\'+'")"}'
            else:
                url = 'http://'+svc_host+':'+svc_port+'/pm_list?where={"myNPWP":"like('+'\\"'+str(exp_npwp)+'\\'+'")","TAHUN_PAJAK":"like('+'\\"'+str(exp_tahun)+'\\'+'")","MASA_PAJAK":"like('+'\\"'+str(exp_masa)+'\\'+'")"}'
            r = requests.get(url,auth = (authresponse['token'],''))
            data = json.loads(r.content)
            if '_meta' in data and data['_meta']['total'] > 25:
                if exp_all == 0:
                    url = 'http://'+svc_host+':'+svc_port+'/pm_list?max_results='+str(data['_meta']['total'])+'&where={"myNPWP":"like('+'\\"'+str(exp_npwp)+'\\'+'")","TAHUN_PAJAK":"like('+'\\"'+str(exp_tahun)+'\\'+'")","MASA_PAJAK":"like('+'\\"'+str(exp_masa)+'\\'+'")","exported":"like('+'\\"'+str(exp_all)+'\\'+'")"}'
                else:
                    url = 'http://'+svc_host+':'+svc_port+'/pm_list?max_results='+str(data['_meta']['total'])+'&where={"myNPWP":"like('+'\\"'+str(exp_npwp)+'\\'+'")","TAHUN_PAJAK":"like('+'\\"'+str(exp_tahun)+'\\'+'")","MASA_PAJAK":"like('+'\\"'+str(exp_masa)+'\\'+'")"}'
                r = requests.get(url,auth = (authresponse['token'],''))
                data = json.loads(r.content)
            if '_items' in data:
                r_items = data['_items']
                for i in r_items:
                    r_list = []
                    r_list.append(str(i['FM']))
                    r_list.append(str(i['KD_JENIS_TRANSAKSI']))
                    r_list.append(str(i['FG_PENGGANTI']))
                    r_list.append(str(i['NOMOR_FAKTUR']))
                    r_list.append(str(i['MASA_PAJAK']))
                    r_list.append(str(i['TAHUN_PAJAK']))
                    tanggalFaktur = i['TANGGAL_FAKTUR']
                    _tanggal, _bulan, _tahun = tanggalFaktur.split('/')                
                    if _tanggal[0] == '0':
                        _tanggal = _tanggal[1]
                    if _bulan[0] == '0':
                        _bulan = _bulan[1]
                    _tanggalFaktur = _tanggal + '/' + _bulan + '/' + _tahun
                    r_list.append(str(_tanggalFaktur))
                    r_list.append(str(i['NPWP']))
                    r_list.append(str(i['NAMA']))
                    r_list.append(str(i['ALAMAT_LENGKAP']))
                    r_list.append(str(i['JUMLAH_DPP']))
                    r_list.append(str(i['JUMLAH_PPN']))
                    r_list.append(str(i['JUMLAH_PPNBM']))
                    r_list.append(str(i['IS_CREDITABLE']))
                    self.data.append(r_list)
                
                ok = True
        except:
            pass
            
        if ok:
            self.signals.expWrite.emit('Writing data...')
            workbook = xlsx_create(exp_filepath)
            worksheet = workbook.add_worksheet()
            
            for r, row in enumerate(self.data):
                self.signals.expPercent.emit(float(r+1)/float(len(self.data))*100)
                for c, col in enumerate(row):
                    worksheet.write(r, c, col)
            workbook.close()
            try:
                url = 'http://'+svc_host+':'+svc_port+'/pm_list?where={"myNPWP":"like('+'\\"'+str(exp_npwp)+'\\'+'")","TAHUN_PAJAK":"like('+'\\"'+str(exp_tahun)+'\\'+'")","MASA_PAJAK":"like('+'\\"'+str(exp_masa)+'\\'+'")"}'
                r = requests.get(url,auth = (authresponse['token'],''))
                data = json.loads(r.content)
                if '_meta' in data and data['_meta']['total'] > 25:
                    url = 'http://'+svc_host+':'+svc_port+'/pm_list?max_results='+str(data['_meta']['total'])+'&where={"myNPWP":"like('+'\\"'+str(exp_npwp)+'\\'+'")","TAHUN_PAJAK":"like('+'\\"'+str(exp_tahun)+'\\'+'")","MASA_PAJAK":"like('+'\\"'+str(exp_masa)+'\\'+'")"}'
                    r = requests.get(url,auth = (authresponse['token'],''))
                    data = json.loads(r.content)
                if '_items' in data:
                    r_items = data['_items']
                    payload = {'exported':1}
                    for r_item in r_items :
                        link = r_item['_links']['self']['href']
                        etag = r_item['_etag']
                
                        url = 'http://'+svc_host+':'+svc_port+'/'+link
                        headers = {'content-type':'application/json','if-match':etag}
                        requests.patch(url,data = json.dumps(payload),headers = headers,auth = (authresponse['token'],''))
            except:
                pass
            self.signals.expFinish.emit(True)
        else:
            self.signals.expLoadData.emit('Failed...')
