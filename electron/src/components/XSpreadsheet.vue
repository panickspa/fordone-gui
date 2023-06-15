<template>
    <div id="spreadsheet" ref="spreadsheet-fordone" style="width:100%; height: 100%; min-height: 100vh;"></div>
</template>

<script lang="ts">
    import Spreadsheet from 'x-data-spreadsheet';
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
            const wrapper:HTMLDivElement|any = document.getElementById('spreadsheet')
            const spreadsheet = new Spreadsheet(wrapper, {
                row:{
                    len: 26,
                    height: 25
                },
                col:{
                    len:9,
                    width: 100,
                    indexWidth: 60,
                    minWidth: 60,
                },
                view: {
                    height: () => wrapper ? wrapper.clientHeight : 0,
                    width: () => wrapper ? wrapper.clientWidth : 0,
                }
            });
            const v = this
            spreadsheet.on('cell-selected', (_cell, ri, ci)=>{
                // v.s = spreadsheet.getData()
                v.$emit('selectedCells', {
                    sci: ci,
                    sri: ri,
                    eci: ci,
                    eri: ri
                });
            })
            spreadsheet.on('cell-edited', ()=> {
                v.s = spreadsheet.getData()
            })
            spreadsheet.on('cells-selected', (_data, location)=>{
                v.$emit('selectedCells', location);
                // console.log({
                //     sci: location.sci,
                //     sri: location.sri,
                //     eci: location.eci,
                //     eri: location.eri
                // })
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