class tImportPM(QtCore.QThread):
    def __init__(self, parent = None):
        QtCore.QThread.__init__(self, parent)
        self.exiting = False
        self.signals = custSignal()
    
    def run(self):
        global importPM_list
        importPM_list = []
        
        if self.exiting == False:
            self.signals.importStart.emit('Reading input file..')
            try:
                wb = xlrd.open_workbook(file2proc)
                data = wb.sheet_by_index(0)
                npwp = data.cell_value(0, 0)
                
                for row in range(2, data.nrows):
                    if data.cell_value(row, 1) == '0':
                        isListed = False
                        for i in xrange(len(importPM_list)):
                            if data.cell_value(row, 1) == importPM_list[i][3] and data.cell_value(row, 2) == importPM_list[i][4]:
                                self.signals.importSkip.emit('Skipped line: ' + str(row+1) + ' | Duplicate Entry  | ' + importPM_list[i][2] + importPM_list[i][3] + importPM_list[i][4])
                                isListed = True
                                break
                        if isListed:
                            continue
                    elif data.cell_value(row, 1) == '1':
                        isListed = False
                        for i in xrange(len(importPM_list)):
                            if str(data.cell_value(row, 5)).strip() == '':
                                self.signals.importSkip.emit('Skipped line: ' + str(row+1) + ' | Contains empty value')
                                isListed = True
                                break
                            tuple_tgl_faktur = xlrd.xldate_as_tuple(data.cell_value(row, 5), wb.datemode)
                            tgl_faktur = str(datetime.date(*tuple_tgl_faktur[:3]))
                            tahun, bulan, tanggal = tgl_faktur.split('-')
                            values = str(tanggal) + '/' + str(bulan) + '/' + str(tahun) 
                            if data.cell_value(row, 1) == importPM_list[i][3] and data.cell_value(row, 2) == importPM_list[i][4] and values == importPM_list[i][7]:
                                self.signals.importSkip.emit('Skipped line: ' + str(row+1) + ' | Duplicate Entry  | ' + importPM_list[i][2] + importPM_list[i][3] + importPM_list[i][4] + ' | ' + importPM_list[i][7])
                                isListed = True
                                break
                        if isListed:
                            continue

                    data_cells = []
                    data_cells.append('1')
                    data_cells.append('')
                    
                    isNull = False
                    tanggalValid = False
                    for col in range(0, data.ncols):
                        if data.cell_type(row, col) == 0:
                            isNull = True
                            break
                        elif data.cell_type(row, col) == 1:
                            values = data.cell_value(row, col).encode('ascii', 'ignore')
                            if values.strip() == '':
                                isNull = True
                                break
                        elif data.cell_type(row, col) == 2:
                            if str(data.cell_value(row, col)).strip() == '':
                                isNull = True
                                break
                            values = int(data.cell_value(row, col))
                        elif data.cell_type(row, col) == 3:
                            if str(data.cell_value(row, col)).strip() == '':
                                isNull = True
                                break
                            tuple_tgl_faktur = xlrd.xldate_as_tuple(data.cell_value(row, col), wb.datemode)
                            tgl_faktur = str(datetime.date(*tuple_tgl_faktur[:3]))
                            tahun, bulan, tanggal = tgl_faktur.split('-')
                            if int(tahun) == int(data.cell_value(row,4)) and int(data.cell_value(row,3)) <= int(bulan)+3 and int(data.cell_value(row,3)) >= int(bulan):
                                tanggalValid = True
                            elif int(data.cell_value(row,4)) == int(tahun) + 1 :
                                if int(data.cell_value(row,3)) < int(bulan):
                                    diff = int(data.cell_value(row,3)) + 12 - int(bulan)
                                    tanggalValid = diff <= 3
                                else :
                                    tanggalValid = False
                            else:
                                tanggalValid = False
                            values = str(tanggal) + '/' + str(bulan) + '/' + str(tahun) 
                        else:
                            values = data.cell_value(row, col)
                        
                        data_cells.append(values)
                    
                    if isNull:
                        self.signals.importSkip.emit('Skipped line: ' + str(row+1) + ' | Contains empty value')
                        continue                    
                    if not tanggalValid:
                        self.signals.importSkip.emit('Skipped line: ' + str(row+1) + ' | tahun atau masa pajak tidak valid')
                        continue
                    data_cells.append(npwp)
                    importPM_list.append(data_cells)
                    
                    _fgPengganti = importPM_list[len(importPM_list)-1][3]
                    _nomorFaktur = importPM_list[len(importPM_list)-1][4]
                    _tanggalFaktur = importPM_list[len(importPM_list)-1][7]
                    if _fgPengganti == '0':
                        url = 'http://'+svc_host+':'+svc_port+'/pm_list?where={"FG_PENGGANTI":"like('+'\\"'+_fgPengganti+'\\'+'")","NOMOR_FAKTUR":"like('+'\\"'+_nomorFaktur+'\\'+'")","myNPWP":"like('+'\\"'+npwp+'\\'+'")"}'
                    else:
                        url = 'http://'+svc_host+':'+svc_port+'/pm_list?where={"FG_PENGGANTI":"like('+'\\"'+_fgPengganti+'\\'+'")","NOMOR_FAKTUR":"like('+'\\"'+_nomorFaktur+'\\'+'")","TANGGAL_FAKTUR":"like('+'\\"'+_tanggalFaktur+'\\'+'")","myNPWP":"like('+'\\"'+npwp+'\\'+'")"}'
                    r = requests.get(url,auth = (authresponse['token'],''))  
                    res = json.loads(r.content)
                    if '_items' in res and len(res['_items']) > 0:
                        importPM_list[len(importPM_list)-1][0] = '0'
                        importPM_list[len(importPM_list)-1][1] = 'Duplicate Entry  | ' + importPM_list[len(importPM_list)-1][2] + _fgPengganti + _nomorFaktur + '  |  ' + res["_items"][0]["TANGGAL_FAKTUR"] + ' (select to replace!)'
                self.signals.importDone.emit('Done')
                if move_processed == 1:
                    prefix = strftime("%Y%m%d_%H%M%S_")
                    new_filename = prefix + os.path.basename(file2proc)
                    new_filepath = os.path.join(str(dir_processed), new_filename)
                    try:
                        if os.path.isfile(new_filepath):
                            os.remove(new_filepath)
                        move(file2proc, new_filepath)
                    except:
                        pass
            except:
                self.signals.importDone.emit('Failed when reading file.. Please check your input file!')
