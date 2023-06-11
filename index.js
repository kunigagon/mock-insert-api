import XLSX from "xlsx";
import moment from "moment";
import axios from "axios";
import { readFile } from 'fs/promises';

const API_SERVICE = 'http://localhost:3006'

const httpClient = axios.create({
    baseURL: `${API_SERVICE}/api-chapanakid/v1/`, //YOUR_API_URL HERE
    headers: {
        "Content-Type": "application/json",
    }
});

const preMapreplace = new Map();
preMapreplace.set('นาย', 'นาย');
preMapreplace.set('พ.ญ.', 'แพทย์หญิง');
preMapreplace.set('นาง', 'นาง');
preMapreplace.set('ร.ต.', 'ร้อยตรี');
preMapreplace.set('น.พ.', 'นายแพทย์');
preMapreplace.set('คุณหญิง', 'คุณหญิง');
preMapreplace.set('จ.ส.อ.', 'จ่าสิบเอก');
preMapreplace.set('ส.อ.', 'สิบเอก');
preMapreplace.set('ว่าที่ ร.ต', 'ว่าที่ร้อยตรี');
preMapreplace.set('ส.ท.', 'สิบโท');
preMapreplace.set('ส.ต.', 'สิบตรี');
preMapreplace.set('จ.ท.', 'จ่าโท');
preMapreplace.set('จ.ท.', 'จ่าโท');
preMapreplace.set('น.ส.', 'นางสาว');
preMapreplace.set('ว่าที่ ร.ท.', 'ว่าที่ร้อยโท');
preMapreplace.set('ว่าที่ ร.ต.', 'ว่าที่ร้อยตรี');
preMapreplace.set('พ.อ.', 'พันเอก');
preMapreplace.set('พ.ต.', 'พันตรี');
preMapreplace.set('มล.', 'หม่อมหลวง');
preMapreplace.set('ด.ต.', 'นายดาบตำรวจ');
preMapreplace.set('ร.อ.', 'ร้อยเอก');
preMapreplace.set('จ.อ.', 'จ่าเอก');
preMapreplace.set('ทญ.', 'ทันตแพทย์หญิง');
preMapreplace.set('จ.ต.', 'จ่าตรี');
preMapreplace.set('ร.ท.', 'ร้อยโท');
preMapreplace.set('ว่าที่ ร.ท', 'ว่าที่ร้อยโท');
preMapreplace.set('ว่าที่ร้อยตรี', 'ว่าที่ร้อยตรี');

const preStatus = new Map();
preStatus.set('ถึงแก่กรรม', 'deceased');
preStatus.set('พ้นสมาชิกภาพ', 'dismiss');
preStatus.set('ลาออก', 'resign');
preStatus.set('ถูกตัดสมาชิก', 'deteriorated')

function sleep(time) {
    return new Promise((resolve) => {
        setTimeout(resolve, time || 1000);
    });
}

function nonEmpty(stringText) {
    return stringText ? stringText : '';
}


const insertMember = async (repairingId = undefined) => {
    console.log('start');
    // const reader = new FileReader();
    // const workbook = XLSX.read(data, { type: "array" });
    // const sheetName = workbook.SheetNames[0];
    // const worksheet = workbook.Sheets[sheetName];

    const fileMember = XLSX.readFile('./tbl_Member-28-04-2566.xlsx');
    const fileMemberRelation = XLSX.readFile('./tbl_Relation.xlsx');
    let dataMember = [];
    let dataMemberRelation = [];

    const sheetsMember = fileMember.SheetNames;
    const sheetsMemberRelation = fileMemberRelation.SheetNames;

    for (let i = 0; i < sheetsMember.length; i++) {
        const temp = XLSX.utils.sheet_to_json(fileMember.Sheets[fileMember.SheetNames[i]]);
        temp.forEach((res) => {
            if (repairingId && repairingId.length > 0) {
                if (repairingId.indexOf(`'${res.CMT_ID}'`) >= 0) {
                    dataMember.push(res)
                }
            } else {
                dataMember.push(res)
            }
        })
    }

    for (let i = 0; i < sheetsMemberRelation.length; i++) {
        const temp = XLSX.utils.sheet_to_json(fileMemberRelation.Sheets[fileMemberRelation.SheetNames[i]]);
        temp.forEach((res) => {
            dataMemberRelation.push(res)
        })
    }

    // console.log('dataMember', dataMember);
    //02596 last update
    const errorList = [];
    let count = 0;
    for (const member of dataMember) {
        if (member.CMT_Name) {
            let preName = '';
            let name = '';
            let lastname = '';
            let keyPrenameOrigin = '';
            if (member.CMT_Sex === 'ชาย') {
                preName = 'นาย'
            } else {
                preName = 'นางสาว';
            }
            for await (const keyPrename of Array.from(preMapreplace.keys())) {
                if (member.CMT_Name.indexOf(keyPrename) > -1) {
                    preName = preMapreplace.get(keyPrename);
                    name = member.CMT_Name.replace(/\s\s+/g, ' ').replace(keyPrename, '').split(' ')[0];
                    lastname = member.CMT_Name.replace(/\s\s+/g, ' ').split(' ')[1];
                }
            }

            console.log('finding prename');
            for await (const keyPrename of Array.from(preMapreplace.keys())) {
                if (member.CMT_Name.indexOf(keyPrename) > -1) {
                    console.log('maped prename');
                    preName = preMapreplace.get(keyPrename);
                    keyPrenameOrigin = keyPrename;
                }
            }

            // console.log('cat prename', keyPrenameOrigin);
            name = member.CMT_Name.replace(/\s\s+/g, ' ').replace(keyPrenameOrigin, '').split(' ')[0];
            lastname = member.CMT_Name.replace(/\s\s+/g, ' ').split(' ')[1];


            // preMap.set(member.CMT_Status_Type, member.CMT_Status_Type);

            const occupationType = member.CMT_Type === 'ข้าราชการ' || member.CMT_Type === 'ลูกจ้างประจำ' || member.CMT_Type === 'พนักงานราชการ' ? 'สามัญ' : 'สมทบ';
            const occupation = member.CMT_Type === 'ข้าราชการ' || member.CMT_Type === 'ลูกจ้างประจำ' || member.CMT_Type === 'พนักงานราชการ' ? member.CMT_Type : '-';
            const occupationOther = member.CMT_Type === 'ข้าราชการ' || member.CMT_Type === 'ลูกจ้างประจำ' || member.CMT_Type === 'พนักงานราชการ' ? null : '-';

            // occupationTypePay
            // <Radio value={'ข้าราชการ'}>ข้าราชการ</Radio>
            // <Radio value={'ลูกจ้างประจำ'}>ลูกจ้างประจำ</Radio>
            // <Radio value={'พนักงานราชการ'}>พนักงานราชการ</Radio>
            // <Radio value={'บำเหน็จ'}>บำเหน็จ</Radio>
            // <Radio value={'บำนาญ'}>บำนาญ</Radio>

            let occupationTypePay = null;

            if (member.CMT_Acc_Name) {
                if (member.CMT_Acc_Name === 'บำเหน็จรายเดือน') {
                    occupationTypePay = 'บำเหน็จ';
                } else if (member.CMT_Acc_Name === 'บำนาญ') {
                    occupationTypePay = 'บำนาญ';
                } else if (member.CMT_Type === 'ข้าราชการ') {
                    occupationTypePay = 'ข้าราชการ';
                } else if (member.CMT_Type === 'ลูกจ้างประจำ') {
                    occupationTypePay = 'ลูกจ้างประจำ';
                } else if (member.CMT_Type === 'พนักงานราชการ') {
                    occupationTypePay = 'พนักงานราชการ';
                }
            }

            const relationAll = await dataMemberRelation.filter((item) => item.RLT_ID === member.CMT_ID);
            let relationManager = await relationAll.filter((item) => item.RLT_Status === 0.4);
            let relationBenefic = await relationAll.filter((item) => item.RLT_Status !== 0.4);
            if (relationManager.length === 0) {
                relationManager = relationBenefic;
            }
            //error list [ '01972', '27947', '41547', '45831', '00000' ]
            let age = getYearDiffWithMonth(new Date(Date.UTC(0, 0, member.CMT_Birthday - 1)), new Date());

            if (age == null || age === null || age < 0) {
                age = 0;
            }
            const memberBody = {
                "occupationType": occupationType,
                "occupationTypePay": occupationTypePay,
                "occupation": occupation,
                "occupationOther": occupationOther,
                "preName": preName,
                "name": name,
                "lastname": lastname,
                "nationalId": member.CMT_IdCard || '-',
                "birthDate": member.CMT_Birthday ? (new Date(Date.UTC(0, 0, member.CMT_Birthday - 1)).toISOString()) : null,
                "age": age,
                "phoneNumber": member.CMT_CPhone ? member.CMT_CPhone.replace(/\s\s+/g, ' ').replace(' ', '').substring(0, 10) : '-',
                "email": `${member.CMT_ID}-dummy@cpk.com`,
                "position": member.CMT_Position || '-',
                "partProject": member.CMT_Office || '-',
                "division": member.CMT_Department || '-',
                "marriedStatus": "-",
                "marriedPreName": "-",
                "marriedName": "-",
                "marriedLastName": "-",
                "marriedNationalId": "-",
                "marriedRegistrationDocument": null,
                "marriedNationIdDocument": null,
                "registrationAddress": `${nonEmpty(member.CMT_Add_No)} ${nonEmpty(member.CMT_Add_Soi)} ${nonEmpty(member.CMT_Add_Road)}, ${nonEmpty(member.CMT_Add_Thambon)}, ${nonEmpty(member.CMT_Add_Amphur)}, ${nonEmpty(member.CMT_Add_Province)}, ${nonEmpty(member.CMT_Add_ZipCode)}`.replace(/\s\s+/g, ' '),
                "address": `${nonEmpty(member.CMT_Add_No)} ${nonEmpty(member.CMT_Add_Soi)} ${nonEmpty(member.CMT_Add_Road)}, ${nonEmpty(member.CMT_Add_Thambon)}, ${nonEmpty(member.CMT_Add_Amphur)}, ${nonEmpty(member.CMT_Add_Province)}, ${nonEmpty(member.CMT_Add_ZipCode)}`.replace(/\s\s+/g, ' '),
                "nationalCardDocument": null,
                "houseRegistrationDocument": null,
                "doctorCertificateDocument": null,
                "manager": relationManager.map((item) => {
                    return {
                        "name": item.RLT_FName ? item.RLT_FName.replace(' ', '') : '-',
                        "lastname": item.RLT_SName ? item.RLT_SName.replace(' ', '') : '-',
                        "nationalId": "-",
                        "phoneNumber": "-",
                        "relation": "-",
                        "relationOther": "-",
                        "address": "-",
                        "managerNationalCardDocument": null,
                        "managerRegistrationDocument": null
                    }
                }),
                "beneficiary": relationBenefic.map((item) => {
                    return {
                        "name": item.RLT_FName ? item.RLT_FName.replace(' ', '') : '-',
                        "lastname": item.RLT_SName ? item.RLT_SName.replace(' ', '') : '-',
                        "nationalId": "-",
                        "phoneNumber": "-",
                        "relation": "-",
                        "relationOther": "-",
                        "address": "-",
                        "managerNationalCardDocument": null,
                        "managerRegistrationDocument": null
                    }
                }),
                "nameRef": null,
                "lastnameRef": null,
                "nationalIdRef": null,
                "memberIdRef": null,


                "accPayById": member.CMT_Acc_Code || null,
                "accPayByName": member.CMT_Acc_Name || null,
                "memberId": member.CMT_ID,
                "status": member.CMT_Status_Type ? preStatus.get(member.CMT_Status_Type) : 'approve',
                "approveDateTime": member.CMT_DateMember ? (new Date(Date.UTC(0, 0, member.CMT_DateMember - 1)).toISOString()) : null,
            };
            if (member.CMT_ID > '00030') {
                // console.log('memberBody', memberBody);
                try {
                    const resUser = await httpClient.post("/users/import", memberBody);
                    if (resUser && resUser.data) {
                        console.log('insert success', `${count} of ${dataMember.length} :`, member.CMT_ID);
                    } else {
                        // console.log('insert error', member.CMT_ID);
                        errorList.push(member.CMT_ID);
                    }
                } catch (e) {
                    // console.log('insert error', e);
                    errorList.push(member.CMT_ID);
                    console.log(member.CMT_ID);
                }
            } else {
                // console.log('skip', member.CMT_ID);
            }



            // console.log('relationManager', relationManager);
            // console.log('relationBenefic', relationBenefic);
            // console.log('key-all', { name: member.CMT_Name, relation: { mname: relationAll.RLT_FName, lastname: relationAll.RLT_SName } });
            count++;
            await sleep(5);
        }
    }

    console.log('end');

    console.log('error list', errorList); s

    // const tx = {
    //     Field1: 28,
    //     CMT_Auto: 1,
    //     CMT_ID: '00001',
    //     CMT_Simple: false,
    //     CMT_Consult: '253201',
    //     CMT_DateConsult: 32793,
    //     CMT_Name: 'นายจริย์  ตุลยานนท์',
    //     CMT_Sex: 'ชาย',
    //     CMT_Birthday: 10606,
    //     CMT_Add_No: '89',
    //     CMT_Add_Soi: 'อัคนี',
    //     CMT_Add_Thambon: 'บางเขน',
    //     CMT_Add_Amphur: 'เมือง',
    //     CMT_Add_Province: 'นนทบุรี',
    //     CMT_Add_ZipCode: '11000',
    //     CMT_Type: 'ข้าราชการ',
    //     CMT_Acc_Code: '987',
    //     CMT_Acc_Name: 'เสียชีวิต',
    //     CMT_Status_Type: 'ถึงแก่กรรม',
    //     CMT_sNumber: 13241,
    //     CMT_Month_Year: 44593,
    //     CMT_Position: 'นักบริหาร 10',
    //     CMT_Office: 'กรมชลประทาน',
    //     CMT_Group: '-',
    //     CMT_Pay: 'กรมบัญชีกลาง',
    //     CMT_Date2: 44553,
    //     CMT_Cause: 'ภาวะหัวใจหยุดเต้นเฉียบพลัน',
    //     CMT_Remark: '60=3คนๆละ 20225.66 17/3/65',
    //     CMT_Balance: 0,
    //     CMT_Debtor: 0,
    //     CMT_Contact: 'นายจริย์  ตุลยานนท์',
    //     CMT_CAdd_No: '89/3',
    //     CMT_CAdd_Soi: 'งามวงศ์วาน 2',
    //     CMT_CAdd_Road: 'งามวงศ์วาน',
    //     CMT_CAdd_Thambon: 'บางเขน',
    //     CMT_CAdd_Amphur: 'เมือง',
    //     CMT_CAdd_Province: 'นนทบุรี',
    //     CMT_CAdd_ZipCode: '11000',
    //     CMT_sMember: 'I',
    //     CMT_cLetter: false
    // };

    // const rt =
    // {
    //     RLT_Auto: 130420,
    //     RLT_ID: '24212',
    //     RLT_FName: 'นายอดิศร',
    //     RLT_SName: 'ดังนางหวาย',
    //     RLT_Reletion: 'บุตร',
    //     RLT_Status: 1,
    //     RLT_Change: 1,
    //     RLT_Date: 44659
    // };


}

const updateDecessRequestMember = async () => {
    console.log('Hello updateDecessRequestMember');
    // const reader = new FileReader();
    // const workbook = XLSX.read(data, { type: "array" });
    // const sheetName = workbook.SheetNames[0];
    // const worksheet = workbook.Sheets[sheetName];

    const fileMember = XLSX.readFile('./tbl_Member-28-04-2566.xlsx');

    let dataMember = [];

    const sheetsMember = fileMember.SheetNames;

    for (let i = 0; i < sheetsMember.length; i++) {
        const temp = XLSX.utils.sheet_to_json(fileMember.Sheets[fileMember.SheetNames[i]]);
        temp.forEach((res) => {
            dataMember.push(res)
        })
    }
    // const deceased = dataMember.filter((item) => item.CMT_Date2 && item.CMT_Status_Type === 'ถึงแก่กรรม' && item.CMT_ID > '41248');
    let deceased = dataMember.filter((item) => (item.CMT_Status_Type === 'ถึงแก่กรรม' || item.CMT_Acc_Code === '987' || item.CMT_Acc_Code === 987 || item.CMT_Acc_Name === 'เสียชีวิต'));
    //(item.CMT_Date2 || item.CMT_Month_Year) && 
    console.log('dataMember', deceased.length);


    const listId = `'00188', '00752', '02836', '03308',
  '03395', '03498', '04110', '04859',
  '05262', '05280', '05764', '06024',
  '06025', '07494', '08609', '09155',
  '09771', '12490', '13070', '13714',
  '15356', '16544', '18313', '18563',
  '18709', '19170', '20444', '20810',
  '21831', '21858', '24450', '25620',
  '26723', '27968', '28777', '28900',
  '29158', '29163', '30839', '31812',
  '32448', '33095', '33703', '34379',
  '36503', '39615', '43649'`;

    const deceasedRepare = await deceased.filter((item) => listId.indexOf(`'${item.CMT_ID}'`) > 0);
    console.log('deceased lenght', deceasedRepare.length);
    const errorList = [];
    let count = 0;
    for (const member of deceasedRepare) {

        let preName = '-';
        let name = '-';
        let lastname = '-';

        // for (const keyPrename of Array.from(preMapreplace.keys())) {
        //     if (member.CMT_Contact && member.CMT_Contact.indexOf(keyPrename) > -1) {
        //         preName = preMapreplace.get(keyPrename);
        //         name = member.CMT_Contact.replace(/\s\s+/g, ' ').replace(keyPrename, '').split(' ')[0];
        //         lastname = member.CMT_Contact.replace(/\s\s+/g, ' ').split(' ')[1];
        //     }
        // }

        const deathRequest = member.CMT_Date1 ? member.CMT_Date1 : member.CMT_Month_Year;
        const deceasedRecord = {
            "refUserId": member.CMT_ID,
            "deathDate": member.CMT_Date2 ? (new Date(Date.UTC(0, 0, member.CMT_Date2 - 1)).toISOString()) : (deathRequest ? (new Date(Date.UTC(0, 0, deathRequest - 1)).toISOString()) : null),
            "deathRequest": deathRequest ? (new Date(Date.UTC(0, 0, deathRequest - 1)).toISOString()) : null,
            "nameRequest": name,
            "lastNameRequest": lastname,
            "nationalIdRequest": "",
            "phoneNumberRequest": member.CMT_CPhone ? member.CMT_CPhone.replace(/\s\s+/g, ' ').replace(' ', '').substring(0, 10) : '-',
            "relationRequest": "",
            "addressRequest": `${nonEmpty(member.CMT_Add_No)} ${nonEmpty(member.CMT_Add_Soi)} ${nonEmpty(member.CMT_Add_Road)}, ${nonEmpty(member.CMT_Add_Thambon)}, ${nonEmpty(member.CMT_Add_Amphur)}, ${nonEmpty(member.CMT_Add_Province)}, ${nonEmpty(member.CMT_Add_ZipCode)}`.replace(/\s\s+/g, ' '),
            "description": `${member.CMT_Assemble || '-'}${member.CMT_Cause || ''}${member.CMT_Remark || ''}`,
            "deceasedDocument1": null,
            "deceasedDocument2": null,
            "deceasedDocument3": null,
            "deceasedDocument4": null,
            "deceasedDocument5": null,
            "deceasedDocument6": null,
            "deceasedDocument7": null,
            "deceasedDocument8": null,
        }

        // const deceasedRecord = {
        //     "refUserId": member.CMT_ID,
        //     "deathDate": new Date('2017-01-01',).toISOString(),
        //     "deathRequest": new Date('2017-01-01').toISOString(),
        //     "nameRequest": name,
        //     "lastNameRequest": lastname,
        //     "nationalIdRequest": "",
        //     "phoneNumberRequest": member.CMT_CPhone ? member.CMT_CPhone.replace(/\s\s+/g, ' ').replace(' ', '').substring(0, 10) : '-',
        //     "relationRequest": "",
        //     "addressRequest": `${nonEmpty(member.CMT_Add_No)} ${nonEmpty(member.CMT_Add_Soi)} ${nonEmpty(member.CMT_Add_Road)}, ${nonEmpty(member.CMT_Add_Thambon)}, ${nonEmpty(member.CMT_Add_Amphur)}, ${nonEmpty(member.CMT_Add_Province)}, ${nonEmpty(member.CMT_Add_ZipCode)}`.replace(/\s\s+/g, ' '),
        //     "description": `${member.CMT_Assemble || '-'}${member.CMT_Cause || ''}${member.CMT_Remark || ''}`,
        //     "deceasedDocument1": null,
        //     "deceasedDocument2": null,
        //     "deceasedDocument3": null,
        //     "deceasedDocument4": null,
        //     "deceasedDocument5": null,
        //     "deceasedDocument6": null,
        //     "deceasedDocument7": null,
        //     "deceasedDocument8": null,
        // }


        console.log('relationManager', JSON.stringify(deceasedRecord));
        // if (member.CMT_ID === '00121') {
        //     console.log('deceasedRecord', `/users/update-deceased/${member.CMT_ID}`, JSON.stringify(deceasedRecord));
        // }

        try {
            const resUser = await httpClient.post(`/users/update-deceased/${member.CMT_ID}`, deceasedRecord);
            if (resUser && resUser.status === 201) {
                console.log('insert success', `${count} of ${deceased.length}`, member.CMT_ID);
            } else {
                console.log('insert error', member.CMT_ID);
                errorList.push(member.CMT_ID);
            }
        } catch (e) {
            console.log('insert error', member.CMT_ID);
            errorList.push(member.CMT_ID);
        }




        // console.log('relationManager', relationManager);
        // console.log('relationBenefic', relationBenefic);
        // console.log('key-all', { name: member.CMT_Name, relation: { mname: relationAll.RLT_FName, lastname: relationAll.RLT_SName } });
        count++;
        await sleep(10);

    }

    console.log('end');

    console.log('error list', errorList);


}

const addZeroPrefix = (number) => {
    // Convert the number to a string
    let str = number.toString();

    // Calculate the number of zeros needed
    let zerosNeeded = 5 - str.length;

    // Add the zeros to the beginning of the string
    let result = '0'.repeat(zerosNeeded) + str;

    return result;
}

const updateDateDecessRequestMember = async () => {
    console.log('Hello updateDateDecessRequestMember');


    const fileMember = XLSX.readFile('./ทั้งหมด.xlsx');

    let dataMember = [];

    const sheetsMember = fileMember.SheetNames;

    for (let i = 0; i < sheetsMember.length; i++) {
        const temp = XLSX.utils.sheet_to_json(fileMember.Sheets[fileMember.SheetNames[i]]);
        temp.forEach((res) => {
            if (res['เลขสมาชิก'] && res['วันที่รับแจ้งเอกสาร']) {
                let memId = addZeroPrefix(res['เลขสมาชิก']);
                dataMember.push({
                    'memberId': memId,
                    'deathRequest': res['วันที่รับแจ้งเอกสาร'] ? (new Date(Date.UTC(0, 0, res['วันที่รับแจ้งเอกสาร'] - 1)).toISOString()) : null
                })
            }
        })
    }

    console.log(dataMember);

    let errorList = [];
    let count = 0;
    for (const item of dataMember) {

        try {
            const resUser = await httpClient.put(`/users/update-deceased-date/${item.memberId}`, item);
            if (resUser && resUser.status) {
                console.log('insert success', `${count} of ${dataMember.length}`, item.memberId);
            } else {
                console.log('insert error', item.memberId);
                errorList.push(item.memberId);
            }
        } catch (e) {
            console.log('insert error', item.memberId);
            errorList.push(item.memberId);
        }

        count++;
        await sleep(10);

    }

    console.log('end');



}


const craateFileupdateDateDecessRequestMember = async () => {
    console.log('Hello craateFileupdateDateDecessRequestMember');


    const fileMember = XLSX.readFile('./รายชื่อคนตายปี2566.xlsx');

    let dataMember = [];

    const sheetsMember = fileMember.SheetNames;

    for (let i = 0; i < sheetsMember.length; i++) {
        const temp = XLSX.utils.sheet_to_json(fileMember.Sheets[fileMember.SheetNames[i]]);
        temp.forEach((res) => {
            if (res['เลขสมาชิก'] && res['วันที่ตาย'] && res['วันที่รับแจ้งเอกสาร']) {
                let memId = addZeroPrefix(res['เลขสมาชิก']);

                const memberId = memId;
                const deathDate = res['วันที่ตาย'] ? (new Date(Date.UTC(0, 0, res['วันที่ตาย'] - 1)).toISOString()) : null;
                const deathRequest = res['วันที่รับแจ้งเอกสาร'] ? (new Date(Date.UTC(0, 0, res['วันที่รับแจ้งเอกสาร'] - 1)).toISOString()) : null;


                console.log(`UPDATE user SET deathRequest = '${deathRequest}' WHERE (memberId = '${memberId}') AND id >0 ;`);
                console.log(`UPDATE record_deceased SET deathDate = '${deathDate}', deathRequest = '${deathRequest}' WHERE id > '0' AND refUserId = (SELECT uuid FROM user where memberId = '${memberId}');`);
                console.log('--');
            }
        })
    }


}


function getYearDiffWithMonth(startDate, endDate) {
    const ms = endDate.getTime() - startDate.getTime();

    const date = new Date(ms);

    return Math.abs(date.getUTCFullYear() - 1970);
}


const updateMissingAcc = async () => {
    console.log('start missing acc');
    const receptAccList = JSON.parse(await readFile(new URL('./recept-acc.json', import.meta.url)));
    // console.log('receptAccList', receptAccList);

    const receptAccListMap = new Map(
        receptAccList.map(object => {
            return [object.code, object.name];
        }),
    );
    const accMisingFile = XLSX.readFile('./acc-missing.xlsx');

    let dataAccUpdate = [];

    const sheetsAccMising = accMisingFile.SheetNames;

    for (let i = 0; i < sheetsAccMising.length; i++) {
        const temp = XLSX.utils.sheet_to_json(accMisingFile.Sheets[accMisingFile.SheetNames[i]]);
        temp.forEach((res) => {
            dataAccUpdate.push({ memberId: res.no, receiptAccId: `${res.accId}`, receiptAccName: receptAccListMap.get(`${res.accId}`) });
        })
    }

    const errorList = [];

    console.log('dataAccUpdate', dataAccUpdate);
    // console.log('dataAccUpdate', dataAccUpdate.filter((item) => item.memberId === '45686' || item.memberId === '45688' || item.memberId === '45824'));
    const amount = dataAccUpdate.length;
    let i = 1;
    for (const accItem of dataAccUpdate) {
        accItem.receiptAccName = 'temp';
        try {
            const resUser = await httpClient.post("/users/update-acc", JSON.stringify(accItem));
            // console.log('resUser', resUser);
            if (resUser && resUser.status === 201) {
                console.log('insert success', `${i} of ${amount} :`, accItem.memberId);
            } else {
                console.log('insert error 1', accItem.memberId);
                errorList.push(accItem.memberId);
            }
        } catch (e) {
            console.log('insert error 2', accItem.memberId);
            errorList.push(accItem.memberId);
        }
        await sleep(10);
        i = i + 1;
    }

    console.log('error list', errorList);
}


// const updateUserApproveDate = async () => {
//     console.log('Hello updateDecessRequestMember');
//     // const reader = new FileReader();
//     // const workbook = XLSX.read(data, { type: "array" });
//     // const sheetName = workbook.SheetNames[0];
//     // const worksheet = workbook.Sheets[sheetName];

//     const fileMember = XLSX.readFile('./tbl_Member.xlsx');

//     let dataMember = [];

//     const sheetsMember = fileMember.SheetNames;

//     for (let i = 0; i < sheetsMember.length; i++) {
//         const temp = XLSX.utils.sheet_to_json(fileMember.Sheets[fileMember.SheetNames[i]]);
//         temp.forEach((res) => {
//             dataMember.push(res)
//         })
//     }
//     const deceased = dataMember.filter((item) => item.CMT_Date2 && item.CMT_Status_Type === 'ถึงแก่กรรม' && item.CMT_ID > '41248');

//     console.log('dataMember', deceased);

//     const errorList = [];
//     let count = 0;
//     for (const member of deceased) {

//         let preName = '';
//         let name = '';
//         let lastname = '';
//         for (const keyPrename of Array.from(preMapreplace.keys())) {
//             if (member.CMT_Contact&&member.CMT_Contact.indexOf(keyPrename) > -1) {
//                 preName = preMapreplace.get(keyPrename);
//                 name = member.CMT_Contact.replace(/\s\s+/g, ' ').replace(keyPrename, '').split(' ')[0];
//                 lastname = member.CMT_Contact.replace(/\s\s+/g, ' ').split(' ')[1];
//             }
//         }
//         // preMap.set(member.CMT_Status_Type, member.CMT_Status_Type);

//         // const occupationType = member.CMT_Type === 'ข้าราชการ' || member.CMT_Type === 'ลูกจ้างประจำ' || member.CMT_Type === 'พนักงานราชการ' ? 'สามัญ' : 'สมทบ';
//         // const occupation = member.CMT_Type === 'ข้าราชการ' || member.CMT_Type === 'ลูกจ้างประจำ' || member.CMT_Type === 'พนักงานราชการ' ? member.CMT_Type : '-';
//         // const occupationOther = member.CMT_Type === 'ข้าราชการ' || member.CMT_Type === 'ลูกจ้างประจำ' || member.CMT_Type === 'พนักงานราชการ' ? null : '-';


//         // const relationAll = await dataMemberRelation.filter((item) => item.RLT_ID === member.CMT_ID);
//         // let relationManager = await relationAll.filter((item) => item.RLT_Status === 0.4);
//         // let relationBenefic = await relationAll.filter((item) => item.RLT_Status !== 0.4);
//         // if (relationManager.length === 0) {
//         //     relationManager = relationBenefic;
//         // }

//         const deceasedRecord = {
//             "refUserId": member.CMT_ID,
//             "deathDate": member.CMT_Date1 ? (new Date(Date.UTC(0, 0, member.CMT_Date1 - 1)).toISOString()) : (member.CMT_Date2 ? (new Date(Date.UTC(0, 0, member.CMT_Date2 - 1)).toISOString()) : null),
//             "deathRequest": member.CMT_Date2 ? (new Date(Date.UTC(0, 0, member.CMT_Date2 - 1)).toISOString()) : null,
//             "nameRequest": name,
//             "lastNameRequest": lastname,
//             "nationalIdRequest": "",
//             "phoneNumberRequest": member.CMT_CPhone ? member.CMT_CPhone.replace(/\s\s+/g, ' ').replace(' ', '').substring(0, 10) : '-',
//             "relationRequest": "",
//             "addressRequest": `${nonEmpty(member.CMT_Add_No)} ${nonEmpty(member.CMT_Add_Soi)} ${nonEmpty(member.CMT_Add_Road)}, ${nonEmpty(member.CMT_Add_Thambon)}, ${nonEmpty(member.CMT_Add_Amphur)}, ${nonEmpty(member.CMT_Add_Province)}, ${nonEmpty(member.CMT_Add_ZipCode)}`.replace(/\s\s+/g, ' '),
//             "description": `${member.CMT_Cause}${member.CMT_Remark ? member.CMT_Remark : ''}`,
//             "deceasedDocument1": null,
//             "deceasedDocument2": null,
//             "deceasedDocument3": null,
//             "deceasedDocument4": null,
//             "deceasedDocument5": null,
//             "deceasedDocument6": null,
//             "deceasedDocument7": null,
//             "deceasedDocument8": null,
//         }

//         // console.log('deceasedRecord', deceasedRecord);
//         try {
//             console.log('calling', member.CMT_ID);
//             const resUser = await httpClient.post(`/users/update-deceased/${member.CMT_ID}`, deceasedRecord);
//             if (resUser) {
//                 console.log('insert success', member.CMT_ID);
//             } else {
//                 console.log('insert error', member.CMT_ID);
//                 errorList.push(member.CMT_ID);
//             }
//         } catch (e) {
//             console.log('insert error', e);
//             errorList.push(member.CMT_ID);
//         }



//         // console.log('relationManager', relationManager);
//         // console.log('relationBenefic', relationBenefic);
//         // console.log('key-all', { name: member.CMT_Name, relation: { mname: relationAll.RLT_FName, lastname: relationAll.RLT_SName } });
//         count++;
//         await sleep(50);

//     }

//     console.log('end');

//     console.log('error list', errorList);


// }

const updatePayingfor = async () => {
    console.log('update paying for');
    //payfor-nov21
    const fileMember = XLSX.readFile('./payfor-nov21.xlsx');

    let dataMember = [];

    const sheetsMember = fileMember.SheetNames;

    for (let i = 0; i < sheetsMember.length; i++) {
        const temp = XLSX.utils.sheet_to_json(fileMember.Sheets[fileMember.SheetNames[i]]);
        temp.forEach((res) => {
            dataMember.push(res)
        })
    }

    for (const member of dataMember) {
        const type0 = `"*01","*02","*03","*04","*05","*06","*07","*08","*09","*10","*11","*12","*13","*14","*15","*16","*17","*4*","*5*","*7*","*8A","*8F","*8I","*8J","*8M","*8N","*8Q","*8R","*9*"`;
        if (type0.indexOf(`"${member.npay}"`) >= 0) {
            console.log(member);
        } else {
            //   console.log("not map");
        }
    }

    console.log('end')

}

const updateType = async (indexSet) => {// ข้าราชการ
    const receptAccList = JSON.parse(await readFile(new URL('./recept-acc.json', import.meta.url)));
    // console.log('receptAccList', receptAccList);

    const receptAccListMap = new Map(
        receptAccList.map(object => {
            return [object.code, object.name];
        }),
    );

    console.log('updateType0');
    const setData = [
        { file: './ข้าราชการ.xlsx', ocupation: 'ข้าราชการ' },
        { file: './บำนาญ.xlsx', ocupation: undefined },
        { file: './บำเหน็จรายเดือน.xlsx', ocupation: undefined },
        { file: './พนักงานราชการ.xlsx', ocupation: 'พนักงานราชการ' },
        { file: './ลูกจ้างประจำ.xlsx', ocupation: 'ลูกจ้างประจำ' }
    ];
    //payfor-nov21
    const fileMember = XLSX.readFile(setData[indexSet].file);

    let dataMember = [];

    const temp = XLSX.utils.sheet_to_json(fileMember.Sheets[fileMember.SheetNames[0]]);
    temp.forEach((res) => {
        dataMember.push(res)
    })
    const errorList = [];

    for (const member of dataMember) {
        let body = {
            memberId: member['CMT_ID'],
            receiptAccId: member['CMT_Acc_Code'],
            receiptAccName: receptAccListMap.get(`${member['CMT_Acc_Code']}`) || 'ไม่พบชื่อสมุห์'
        }
        if (setData[indexSet].ocupation) {
            body['occupation'] = setData[indexSet].ocupation;
        }

        if (member['CMT_Type'] && member['CMT_Type'].indexOf('หักแทนกัน') >= 0) {
            body['memberNationalIdPayInstead'] = `${member['Pay_ID']}`;
        }

        if (!body.receiptAccName || !body.receiptAccId) {
            console.log(body);
        }

        try {
            const resUser = await httpClient.post("/users/update-acc-with-ref", body);
            if (resUser && resUser.status === 201) {
                console.log('insert success', body.memberId);
            } else {
                console.log('insert error 1', body.memberId);
                errorList.push(body.memberId);
            }
        } catch (e) {
            console.log('insert error 2', e);
            errorList.push(body.memberId);
        }
        await sleep(20);
    }

    console.log('end :', errorList.length);
    // console.log('error list', errorList);
    return errorList;
}


const updatePaymentGroup = async () => {
    console.log('start updatePaymentGroup');

    const accMisingFile = XLSX.readFile('./หมวดการหักเงิน.xlsx');

    let dataPaymentGroup = [];

    const sheetsAccMising = accMisingFile.SheetNames;

    for (let i = 0; i < sheetsAccMising.length; i++) {
        const temp = XLSX.utils.sheet_to_json(accMisingFile.Sheets[accMisingFile.SheetNames[i]]);
        temp.forEach((res) => {
            // if (res['หมวด'] > 5)
            dataPaymentGroup.push(
                {
                    memberId: res['เลขสมาชิก'],
                    groupType: res['หมวด'],
                    nationalIdPayInstead: res['เลขบัตรประชาชนที่ใช้เรียกเก็บ']
                }
            )
        })
    }

    console.log('dataPaymentGroup', dataPaymentGroup);

    const errorList = [];
    let i = 0;
    for (const member of dataPaymentGroup) {

        try {
            const resUser = await httpClient.post("/utils/payment-group", member);
            if (resUser && resUser.status === 201) {
                console.log('insert success', `${i} of ${dataPaymentGroup.length}`, member.memberId);
            } else {
                console.log('insert error 1', member.memberId);
                errorList.push(member.memberId);
            }
        } catch (e) {
            console.log('insert error 2', e);
            errorList.push(member.memberId);
        }
        await sleep(10);
        i = i + 1;
    }

    console.log('end :', errorList.length);

}
const main = async () => {
    // await insertMember(undefined);
    // await updateMissingAcc();
    // await updateDecessRequestMember();
    // await updatePayingfor();
    // const res0 = await updateType(0);
    // const res1 = await updateType(1);
    // const res2 = await updateType(2);
    // const res3 = await updateType(3);
    // const res4 = await updateType(4);


    // console.log('res0', res0);
    // console.log('res1', res1);
    // console.log('res2', res2);
    // console.log('res3', res3);
    // console.log('res4', res4);


    // await updatePaymentGroup();

    // await updateDateDecessRequestMember();


    await craateFileupdateDateDecessRequestMember();



}

main();