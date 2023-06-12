<template>
  <v-responsive class="align-center text-center fill-height">
    <v-row>
      <v-col cols="6">
        <v-card width="50vw">
          <v-card-text>
            <x-spreadsheet @selectedCell="$event => selectedCell($event)" @selectedCells="setRanges"></x-spreadsheet>
          </v-card-text>
        </v-card>
      </v-col>
      <v-col cols="6">
        <v-btn @click="getHtml">Get Data</v-btn>
        <v-row>
          <v-col>
            <p>{{ selectedCells }}</p>
            <p>{{ ranges }}</p>
          </v-col>
        </v-row>
      </v-col >
    </v-row>
  </v-responsive>
</template>

<script lang="ts">
  import { cp } from 'fs';
import XSpreadsheet from './XSpreadsheet.vue';
  import Record, {
    SheetData,
  } from 'x-data-spreadsheet'
  import {
    utils,
    CellObject,
    Sheet
  } from 'xlsx'
  // import xtos from 'xlsx';
  export default{
    components:{
      'x-spreadsheet': XSpreadsheet
    },
    methods:{
      getHtml(){
        this.sheet = this.selectedCells
      },
      setRanges(e:{
        sri:number,
        sci:number,
        eri:number,
        eci:number
      }){
        this.ranges = e
      },
      selectedCell(e:any){
        // this.selectedCells = e
        let data = e[0]
        const r:Array<any> = []
        if(data)
          // if(data.rows)
          for(let i=this.ranges.sri; i<this.ranges.eri+1; i++){
            let row:Array<any> = []
            // let col:Array<any> = []
            for(let j=this.ranges.sci; j<this.ranges.eci+1; j++){
              if(data.rows[i]){
                if(data.rows[i].cells[j])
                  row.push(data.rows[i].cells[j] ? data.rows[i].cells[j].text : '' )
              }
              else row.push('')
              // else j = data.cols.len
            }
            r.push(row)
          }
        this.selectedCells = r
        // this.selectedCells = r
        // let s:Sheet = utils.aoa_to_sheet(
        //   r
        // )
        // s['!merges'] = e.merges

        // s['!merges']= e.merges
        
        // s['!data']
        
        // s['!data'] = e.map(e => new CellObj)
        // let b = utils.book_new()
        // console.log(b)
        // utils.book_append_sheet(b, utils.json_to_sheet(e))
        // b.SheetNames
        // this.sheet = utils.sheet_to_html(utils.json_to_sheet(e))
        // this.sheet = utils.format_cell()
        // utils.
        // utils.
        // this.selectedCells = utils.sheet_to_html(b.Sheets[0])
      }
    },
    mounted(){
      
    },
    data(){
      let selectedCells:any
      return{
        selectedCells: selectedCells,
        sheet: selectedCells,
        ranges: {
          sri: 0,
          sci: 0,
          eri: 0,
          eci: 0
        },
      }
    }
  }
</script>
