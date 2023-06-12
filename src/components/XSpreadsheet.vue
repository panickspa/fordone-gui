<template>
    <div id="spreadsheet" ref="spreadsheet-fordone" style="width:100%; max-width: 50vw;">
    </div>
</template>

<script lang="ts">
    import Spreadsheet, { Cell } from 'x-data-spreadsheet';
    export default {
        methods:{
            selectedCell(cell:any, range:any){
                this.$emit('cells-selected', {
                    cell: cell,
                    range: range
                })
            }
        },     
        mounted(){
            const wrapper:any = document.getElementById('spreadsheet')
            const spreadsheet = new Spreadsheet(wrapper);
            const v = this
            spreadsheet.on('cell-selected', ()=>{
                v.s = spreadsheet.getData()
            })
            spreadsheet.on('cell-edited', ()=> {
                v.s = spreadsheet.getData()
            })
            spreadsheet.on('cells-selected', (data, location)=>{
                v.$emit('selectedCells', {
                    sci: location.sci,
                    sri: location.sri,
                    eci: location.eci,
                    eri: location.eri
                } );
                console.log({
                    sci: location.sci,
                    sri: location.sri,
                    eci: location.eci,
                    eri: location.eri
                })
                v.s = spreadsheet.getData()
            })
            this.$emit('selectedCell', spreadsheet.getData())
            // this.s = spreadsheet
        },
        data(){
            let s:any
            return {
                s: s
            }
        },
        watch:{
            s(s){
                // console.log(s.param)
                this.$emit('selectedCell', s)
            }
        }
    }
</script>

<style scoped>

</style>