import fs from 'fs-extra'
import path from 'path'

fs.emptyDir('./sendList/', (err) => {
    console.log(err)
})