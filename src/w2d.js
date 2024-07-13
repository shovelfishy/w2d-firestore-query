import { initializeApp } from "firebase/app";
import { Firestore, getFirestore, query, where, collection, getDocs, or, documentId, orderBy} from "firebase/firestore";
import * as xlsx from 'xlsx'; 


const firebaseConfig = {
    apiKey: "AIzaSyDbupSgO0V8gxquPJqaMGQZ9rzY7oXEoA8",
    authDomain: "w2dapp.firebaseapp.com",
    projectId: "w2dapp",
    storageBucket: "w2dapp.appspot.com",
    messagingSenderId: "1081340134333",
    appId: "1:1081340134333:web:0fd1421a8c11d7faa708d4",
    measurementId: "G-L9GXGXS9CN"
};

const app = initializeApp(firebaseConfig);
const db = getFirestore();
const col = collection(db, 'emailList');

export const fetchAOO = () => {
    const orderedByTime = query(col, orderBy('subscribedAt'));
    return getDocs(orderedByTime)
    .then((snapshot) => {
        let data = [];
        snapshot.docs.forEach((doc)=>{
            const document = doc.data(); 
            data.push({...document, subscribedAt: document['subscribedAt'].toDate().toString(), id: doc.id});
        })
        return data;
    });
}

const excel = async () => {
    var workbook = xlsx.utils.book_new();
    var worksheet = xlsx.utils.json_to_sheet(await fetchAOO());
    worksheet['!cols'] = [ { wch: 32}, {wch: 51}, {wch: 28 } ]
    xlsx.utils.book_append_sheet(workbook, worksheet, "Email List");
    xlsx.writeFile(workbook, 'w2d-emailList.xlsx');
}

await excel();e
console.log('Document successfully created');