<template>
  <v-responsive height="100vh">
    <v-row>
      <v-col cols="6">
        <v-card width="50vw">
          <v-card-text>
            <x-spreadsheet @selectedCell="$event => selectedCell($event)" @selectedCells="setRanges"></x-spreadsheet>
          </v-card-text>
        </v-card>
      </v-col>
      <v-col cols="6">
        <v-btn @click="getHeader" class="mr-2">Set Header</v-btn>
        <v-btn @click="getData">Set Data</v-btn>
        <v-row>
          <v-col>
            <!-- <p>{{ selectedCells }}</p> -->
            <!-- <p>{{ rangesStr }}</p> -->
            <!-- <p>{{ ranges }}</p> -->
            <!-- <p>{{ sheet }}</p> -->
            <!-- <p>{{ sheetHead }}</p> -->
            <!-- <p>{{ aoa }}</p> -->
            <!-- <p>{{ sheet }}</p> -->
            <!-- <p>{{ merges }}</p> -->
            <!-- <p>{{ mergesRaw }}</p> -->
            <!-- <p>{{ msg.err }}</p> -->
          </v-col>
        </v-row>
        <v-row>
          <v-col cols="12">
            <v-responsive height="100%">
              <v-textarea v-model="sheetTable"></v-textarea>
            </v-responsive>
          </v-col>
        </v-row>
      </v-col >
    </v-row>
  </v-responsive>
    
</template>

<script lang="ts">
  import { cp } from 'fs';
  import XSpreadsheet from './XSpreadsheet.vue';
  // import pretty from 'pretty'
  import Record, {
    SheetData,
  } from 'x-data-spreadsheet'
  import xmlFormat from 'xml-formatter'
  import {
    utils,
    CellObject,
    Sheet
  } from 'xlsx'
  // import xtos from 'xlsx';
  interface range {
      sri: number,
      sci: number,
      eri: number,
      eci: number
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
    computed:{
      // selectedCells(){
        
      // },
    },
    methods:{
      getHeader(){
        this.sheetHead =  `<thead>${this.getHtml().replaceAll(/<table>|<\/table>/g,'').replaceAll(/td/g, 'th')}</thead>`
      },
      getData(){
        this.sheetData = `<tbody>${this.getHtml().replaceAll(/<table>|<\/table>/g,'')}</tbody>`
      },
      getHtml():string{
        const r:Array<any> = []
        let row:Array<number|string> = []
        const merges:Array<xlsxRange> = this.sheet
          .merges
          .map((m:string) => utils.decode_range(m))
          .filter((m:xlsxRange) => m.s.r >= this.ranges.sri && m.e.r <= this.ranges.eri && m.s.c >= this.ranges.sci && m.e.c <= this.ranges.eci)
          .map((m:xlsxRange):xlsxRange => (
            {
              s:{
                r: m.s.r - this.ranges.sri,
                c: m.s.c - this.ranges.sci
              },
              e:{
                r: m.e.r - this.ranges.sri,
                c: m.e.c - this.ranges.sci
              }
            }
          ))
        this.mergesRaw = merges.map((e:xlsxRange) => utils.encode_range(e.s,e.e))
        if(this.sheet)
          // if(data.rows)
          for(let i=this.ranges.sri; i<this.ranges.eri+1; i++){
            row = []
            // let col:Array<any> = []
            for(let j=this.ranges.sci; j<this.ranges.eci+1; j++){
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
          if(this.ranges.eri > -1)
            this.aoa = r
            if(r.length > 0){
              let sheet = utils.aoa_to_sheet(r)
              sheet['!merges'] = merges
              return utils.sheet_to_html(sheet,{
                header: '',
                footer: '',
              })
            }
          return ''
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
      }
    },
    data(){
      let selectedCells:any
      return{
        // selectedCells: selectedCells,
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
        aoa:selectedCells,
        mergesRaw: [''],
        msg:{
          err: selectedCells
        },
        sheetTable:''
      }
    },
    watch:{
      sheetData(n:string){
        this.sheetTable = xmlFormat(`<table>${this.sheetHead}
        ${n}</table>`, {
          indentation: '    '
        })
      },
      sheetHead(n:string){
        this.sheetTable = xmlFormat(`<table>${n}
        ${this.sheetData}</table>`,{
          indentation: '    '
        })
      }
    }
  }
</script>
