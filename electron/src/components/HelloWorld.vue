<template>
  <v-responsive height="100vh">
    <v-row>
      <v-col cols="6">
        <v-card width="100%" flat>
          <x-spreadsheet @selectedCell="$event => selectedCell($event)" @selectedCells="setRanges"></x-spreadsheet>
        </v-card>
      </v-col>
      <v-col cols="6">
        <v-card flat class="px-3 py-3">
          <v-row class="mt-3">
            <v-col :cols="3">
              <v-btn @click="getHeader" class="mb-2 w-100">Set Header</v-btn>
              <v-text-field variant="solo" label="Header" v-model="sheetHead"></v-text-field>
            </v-col>
            <v-col :cols="3">
              <v-btn @click="getNumberHeader" class="mb-2 w-100">Set Number Header</v-btn>
              <v-text-field variant="solo" label="Number Header" v-model="sheetNumberHeader"></v-text-field>
            </v-col>
            <v-col :cols="3">
              <v-btn @click="getData" class="mb-2 w-100">Set Fix Data</v-btn>
              <v-text-field variant="solo" label="Fixed Data" v-model="sheetData"></v-text-field>
            </v-col>
            <v-col :cols="3">
              <v-btn @click="getEditedData" class="mb-2 w-100">Set Edited Data</v-btn>
              <v-text-field variant="solo" label="Edited Data" v-model="sheetEditedData"></v-text-field>
            </v-col>
          </v-row>
          <v-row class="mt-3">
            <!-- <v-col cols="12">
              <p>{{ ranges }}</p>
            </v-col> -->
            <v-col cols="12">
              <v-textarea v-model="sheetTable" variant="solo" label="HTML Table"  bg-color="amber-lighten-4" no-resize color="black" rows="20"></v-textarea>
            </v-col>
            <v-col cols="12">
              <v-btn class="m-3" @click="getHtml">Get HTML</v-btn>
            </v-col>
          </v-row>
        </v-card>
      </v-col >
    </v-row>
  </v-responsive>
</template>

<script lang="ts">
  import XSpreadsheet from './XSpreadsheet.vue';
  // import pretty from 'pretty'
  import xmlFormat from 'xml-formatter'
  import {
Range,
    utils,
  } from 'xlsx'
  // import xtos from 'xlsx';
  interface range {
    sci: number,
    sri: number,
    eci: number,
    eri: number
  }
  interface xlsxRange {
    s:{
      r:number,
      c:number
    },
    e:{
      r:number,
      c:number
    }
  }
  export default{
    components:{
      'x-spreadsheet': XSpreadsheet
    },
    methods:{
      getHeader(){
        this.sheetHead = this.encodeRange(this.ranges)
        // this.sheetHead =  this.getHtml().replaceAll(/<table>|<\/table>/g,'').replaceAll(/td/g, 'th')
      },
      getNumberHeader(){
        this.sheetNumberHeader = this.encodeRange(this.ranges)
      },
      getData(){
        this.sheetData = this.encodeRange(this.ranges)
      },
      getEditedData(){
        this.sheetEditedData = this.encodeRange(this.ranges)
      },
      // getStringHtml(){
      //   this.sheetTable = this.getHtml()
      // },
      inRanges(
        e:{
          r:number,
          c:number,
          range:xlsxRange
        }
      ){
        return e.range.s.r <= e.r && e.range.s.c <= e.c && e.range.e.c >= e.c && e.range.e.r >= e.r
      },
      getHtml(){
        let r:Array<any> = []
        let row:Array<number|string> = []

        let h:string = ''
        let n:string = ''
        let d:string = ''

        const noRange:xlsxRange = {
          s:{
            r: -1,
            c: -1,
          },
          e:{
            r: -1,
            c: -1
          }
        }

        let hr:Range|xlsxRange = this.sheetHead ? utils.decode_range(this.sheetHead) : noRange
        let nr:Range|xlsxRange = this.sheetNumberHeader ? utils.decode_range(this.sheetNumberHeader) : noRange
        let fr:Range|xlsxRange = this.sheetData ? utils.decode_range(this.sheetData) : noRange
        let dr:Range|xlsxRange = this.sheetEditedData ? utils.decode_range(this.sheetEditedData) : noRange

        if(this.sheet){
          // Iteration of header
          if(hr !== noRange){
            r = []
            for(let i=hr.s.r; i<hr.e.r+1; i++){
              row = []
              // let col:Array<any> = []
              for(let j=hr.s.c; j<hr.e.c+1; j++){
                if(this.sheet.rows[i]){
                  if(this.sheet.rows[i].cells[j]){
                    row.push(this.sheet.rows[i].cells[j] ? this.sheet.rows[i].cells[j].text : '' )
                    // m.s.r = i
                  }
                  else row.push('')
                }
                else row.push('')
                // else j = data.cols.len
              }
              r.push(row)
            }
            let sheet = utils.aoa_to_sheet(r)
            sheet['!merges'] = this.sheet
                .merges
                .map((m:string) => utils.decode_range(m))
                .filter((m:xlsxRange) => m.s.r >= hr.s.r && m.e.r <= hr.e.r && m.s.c >= hr.s.c && m.e.c <= hr.e.c)
                .map((m:xlsxRange):xlsxRange => (
                  {
                    s:{
                      r: m.s.r - hr.s.r,
                      c: m.s.c - hr.s.c
                    },
                    e:{
                      r: m.e.r - hr.s.r,
                      c: m.e.c - hr.s.c
                    }
                  }
                ))
            h = utils.sheet_to_html(sheet,{
              header: '',
              footer: '',
            }).replaceAll(/<table>|<\/table>/g,'').replaceAll(/td/g, 'th')
              // .replaceAll(/data-t="\w+"\s|data-v="(\$edit(\w|\d|\s)\$end-edit)?|\w?"\s|id="sjs-\w+"/g, '')
          }
          if(nr !== noRange){
            r = []
            for(let i=nr.s.r; i<nr.e.r+1; i++){
              row = []
              // let col:Array<any> = []
              for(let j=nr.s.c; j<nr.e.c+1; j++){
                if(this.sheet.rows[i]){
                  if(this.sheet.rows[i].cells[j]){
                    row.push(this.sheet.rows[i].cells[j] ? this.sheet.rows[i].cells[j].text : '' )
                    // m.s.r = i
                  }
                  else row.push('')
                }
                else row.push('')
                // else j = data.cols.len
              }
              r.push(row)
            }
            let sheet = utils.aoa_to_sheet(r)
            sheet['!merges'] = this.sheet
                .merges
                .map((m:string) => utils.decode_range(m))
                .filter((m:xlsxRange) => m.s.r >= nr.s.r && m.e.r <= nr.e.r && m.s.c >= nr.s.c && m.e.c <= nr.e.c)
                .map((m:xlsxRange):xlsxRange => (
                  {
                    s:{
                      r: m.s.r - nr.s.r,
                      c: m.s.c - nr.s.c
                    },
                    e:{
                      r: m.e.r - nr.s.r,
                      c: m.e.c - nr.s.c
                    }
                  }
                ))
            n = utils.sheet_to_html(sheet,{
              header: '',
              footer: '',
            }).replaceAll(/<table>|<\/table>/g,'')
              // .replaceAll(/data-t="\w+"\s|data-v="(\$edit(\w|\d|\s)\$end-edit)?|\w?"\s|id="sjs-\w+"/g,'')
          }
          if(fr !== noRange && fr !== noRange){
            r = []
            for(let i=fr.s.r; i<dr.e.r+1; i++){
              row = []
              // let col:Array<any> = []
              for(let j=fr.s.c; j<dr.e.c+1; j++){
                if(this.sheet.rows[i]){
                  if(this.sheet.rows[i].cells[j]){
                    row.push(this.sheet.rows[i].cells[j] ? this.inRanges({
                      r: i,
                      c: j,
                      range: fr
                    }) ? this.sheet.rows[i].cells[j].text : `$edit${this.sheet.rows[i].cells[j].text}$end-edit` : this.inRanges({
                      r: i,
                      c: j,
                      range: dr
                    }) ? `$edit $end-edit` : '' )
                    // m.s.r = i
                  }
                  else row.push(
                    this.inRanges({
                      r: i,
                      c: j,
                      range: dr
                    }) ? `$edit $end-edit` : ''
                  )
                }
                else row.push(
                  this.inRanges({
                      r: i,
                      c: j,
                      range: dr
                    }) ? `$edit $end-edit` : ''
                )
                // else j = data.cols.len
              }
              r.push(row)
            }
            let sheet = utils.aoa_to_sheet(r)
            sheet['!merges'] = this.sheet
                .merges
                .map((m:string) => utils.decode_range(m))
                .filter((m:xlsxRange) => m.s.r >= fr.s.r && m.e.r <= dr.e.r && m.s.c >= fr.s.c && m.e.c <= dr.e.c)
                .map((m:xlsxRange):xlsxRange => (
                  {
                    s:{
                      r: m.s.r - fr.s.r,
                      c: m.s.c - fr.s.c
                    },
                    e:{
                      r: m.e.r - fr.s.r,
                      c: m.e.c - fr.s.c
                    }
                  }
                ))
            d = utils.sheet_to_html(sheet,{
              header: '',
              footer: '',
            })
            .replaceAll(/data-t="\w+"\s|data-v="(\$edit(\w|\d|\s)\$end-edit)?"\s|id="sjs-\w+"/g, '')
            .replaceAll(/<table>|<\/table>/g,'')
            .replaceAll(/\$edit/g, `<div class="form-control" id="textarea_tabllewis" contenteditable="">`)
            .replaceAll(/\$end-edit/g, "</div>")
          }
        }
        this.sheetTable = xmlFormat(`<table id="lewistabel"><thead>${h}</thead><tbody>${n}${d}</tbody></table>`,{
          indentation: `  `
        })
        // this.sheetTable = xmlFormat(`<table id="lewistabel">${h}${n}${d}</table>`, {
        //   indentation: `  `
        // })
        // else  return ''
      },
      setRanges(e:{sri:number,sci:number,eri:number,eci:number}){
        if(typeof e == 'string') this.rangesStr = e
        else this.rangesStr = JSON.stringify(e)
        // let r = utils.decode_range(e)
        this.ranges = {
          sri: e.sri,
          sci: e.sci,
          eri: e.eri,
          eci: e.eci
        }
      },
      selectedCell(e:any){
        this.sheet = e[0]
      },
      encodeRange(range:range):string{
        if(range){
          return utils.encode_range({
            s:{
              r:range.sri,
              c:range.sci
            },
            e:{
              r:range.eri,
              c:range.eci
            }
          })
        }
        return ''
      }
    },
    data(){
      let selectedCells:any
      return{
        sheet: selectedCells,
        rangesStr: selectedCells,
        sheetHead: '',
        ranges: {
          sri: 0,
          sci: 0,
          eri: 0,
          eci: 0
        },
        sheetData: '',
        merges: [],
        sheetEditedData: '',
        sheetNumberHeader: '',
        htmlData: '',
        htmlHeader: '',
        htmlNumberHeader: '',
        aoa:selectedCells,
        mergesRaw: [''],
        msg:{
          err: selectedCells
        },
        sheetTable:''
      }
    },
    watch:{
      // sheetData(n:string){
      //   this.sheetTable = xmlFormat(`<table>${this.sheetHead}
      //   ${n}</table>`, {
      //     indentation: '    '
      //   })
      // },
      // sheetHead(n:string){
      //   this.sheetTable = xmlFormat(`<table>${n}
      //   ${this.sheetData}</table>`,{
      //     indentation: '    '
      //   })
      // }
    }
  }
</script>
