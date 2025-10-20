/**
 * Mapping of day abbreviations to day indices
 * @type {Object.<string, number>}
 */
const dayIndex = {
  'Sun.': 0,
  'M.': 1,
  'T.': 2,
  'W.': 3,
  'Th.': 4,
  'F.': 5,
  'Sat.': 6
};

const CONFIG = {
  APP: {
    NAME: 'Paybox - DPU Monitoring',
    VERSION: '2.1.0',
    ADMIN:{
      name: 'Erwin Alcantara',
      email: 'egalcantara@multisyscorp.com'
    }
  },

  ETAP: {
    SERVICE_BANK: 'eTap',
    ENVIRONMENT: 'production',
    EMAIL_RECIPIENTS: {
      production: {
        to: "christian@etapinc.com, jayr@etapinc.com, rodison.carapatan@etapinc.com, arnel@etapinc.com, ian.estrella@etapinc.com, roldan@etapinc.com, dante@etapinc.com, ramun.hamtig@etapinc.com, jellymae.osorio@etapinc.com, rojane@etapinc.com, A.jloro@etapinc.com, richard@etapinc.com, johnrandy.divina@etapinc.com",
        cc: "reinier@etapinc.com, miguel@etapinc.com, alvie@etapinc.com, rojane@etapinc.com, laila@etapinc.com, johnmarco@etapinc.com, ghie@etapinc.com, etap-recon@etapinc.com, Erwin Alcantara <egalcantara@multisyscorp.com>, cvcabanilla@multisyscorp.com, gmmrectin@multisyscorp.com, amlaurio@multisyscorp.com",
        bcc: "mvolbara@pldt.com.ph, RBEspayos@smart.com.ph"
      },
      testing: {
        to: "Erwin Alcantara <egalcantara@multisyscorp.com>",
        cc: "Erwin Alcantara <egalcantara@multisyscorp.com>",
        bcc: ""
      }
    },
    SPECIAL_MACHINES: new Set(['PLDT BANTAY', 'SMART VIGAN'])
  },

  BRINKS: {
    SERVICE_BANK: 'Brinks via BPI',
    ENVIRONMENT: 'production',
    EMAIL_RECIPIENTS: {
      production: {
        to: "mbocampo@bpi.com.ph",
        cc: "julsales@bpi.com.ph, mjdagasuhan@bpi.com.ph, eagmarayag@bpi.com.ph,rtorticio@bpi.com.ph, egcameros@bpi.com.ph, vjdtorcuator@bpi.com.ph, jdduque@bpi.com.ph, rsmendoza1@bpi.com.ph,jmdcantorna@bpi.com.ph, kdrepuyan@bpi.com.ph, dabayaua@bpi.com.ph, rdtayag@bpi.com.ph, vrvarellano@bpi.com.ph,mapcabela@bpi.com.ph, mvpenisa@bpi.com.ph, mbcernal@bpi.com.ph, cmmanalac@bpi.com.ph, mpdcastro@bpi.com.ph,rmdavid@bpi.com.ph, emflores@bpi.com.ph, apmlubaton@bpi.com.ph, smcarvajal@bpi.com.ph, avabarabar@bpi.com.ph,jcmontes@bpi.com.ph, jeobautista@bpi.com.ph, micaneda@bpi.com.ph, rrpacio@bpi.com.ph,mecdycueco@bpi.com.ph, tesruelo@bpi.com.ph, ssibon@bpi.com.ph, christine.sarong@brinks.com, icom2.ph@brinks.com,aillen.waje@brinks.com, rpsantiago@bpi.com.ph, jerome.apora@brinks.com, occ2supervisors.ph@brinks.com, mdtenido@bpi.com.ph, agmaiquez@bpi.com.ph, jsdamaolao@bpi.com.ph, cvcabanilla@multisyscorp.com, raflorentino@multisyscorp.com, egalcantara@multisyscorp.com, gmmrectin@multisyscorp.com, amlaurio@multisyscorp.com",
        bcc: "mvolbara@pldt.com.ph, RBEspayos@smart.com.ph",
      },
      testing: {
        to: "Erwin Alcantara <egalcantara@multisyscorp.com>",
        cc: "Erwin Alcantara <egalcantara@multisyscorp.com>",
        bcc: "",
      },
    },
    SPECIAL_MACHINES: new Set([])
  },

  BPI_INTERNAL: {
    SERVICE_BANK: 'BPI Internal',
    ENVIRONMENT: 'production',
    EMAIL_RECIPIENTS: {
      production: {
        to: "mjdagasuhan@bpi.com.ph",
        cc: "malodriga@bpi.com.ph, julsales@bpi.com.ph, jcslingan@bpi.com.ph, eagmarayag@bpi.com.ph, jmdcantorna@bpi.com.ph, egcameros@bpi.com.ph, jdduque@bpi.com.ph, rdtayag@bpi.com.ph, rsmendoza1@bpi.com.ph, vjdtorcuator@bpi.com.ph, egalcantara@multisyscorp.com, raflorentino@multisyscorp.com, cvcabanilla@multisyscorp.com, gmmrectin@multisyscorp.com, amlaurio@multisyscorp.com",
        bcc: "mvolbara@pldt.com.ph, RBEspayos@smart.com.ph"
      },
      testing: {
        to: "Erwin Alcantara <egalcantara@multisyscorp.com>",
        cc: "Erwin Alcantara <egalcantara@multisyscorp.com>",
        bcc: ""
      },
    },
    SPECIAL_MACHINES: new Set([])
  },

  APEIROS: {
    SERVICE_BANK: 'Apeiros',
    ENVIRONMENT: 'production',  //testing
    EMAIL_RECIPIENTS: {
      production: {
        // to: "mtcsantiago570@gmail.com, mtcsurigao@gmail.com, valdez.ezekiel23@gmail.com",
        // cc: "sherwinamerica@yahoo.com, egalcantara@multisyscorp.com, raflorentino@multisyscorp.com, cvcabanilla@multisyscorp.com, gmmrectin@multisyscorp.com, amlaurio@multisyscorp.com",
        // bcc: "mvolbara@pldt.com.ph"
        to: "Erwin Alcantara <egalcantara@multisyscorp.com>",
        cc: "Erwin Alcantara <egalcantara@multisyscorp.com>",
        bcc: ""
      },
      testing: {
        to: "Erwin Alcantara <egalcantara@multisyscorp.com>",
        cc: "Erwin Alcantara <egalcantara@multisyscorp.com>",
        bcc: ""
      },
    },
    SPECIAL_MACHINES: new Set(['PLDT SANTIAGO','SMART ROBINSONS SANTIAGO'])
  },

  SCHEDULE_BASED: {
    STORES: ["PLDT ROBINSONS DUMAGUETE", "SMART SM BACOLOD 1", "SMART SM BACOLOD 2", "SMART SM BACOLOD 3", "PLDT ILIGAN", "SMART GAISANO MALL OZAMIZ"],
    SCHEDULES: {
      "PLDT ROBINSONS DUMAGUETE": [dayIndex["M."], dayIndex["W."], dayIndex["Sat."]],
      "SMART SM BACOLOD 1": [dayIndex["T."], dayIndex["Sat."]],
      "SMART SM BACOLOD 2": [dayIndex["T."], dayIndex["Sat."]],
      "SMART SM BACOLOD 3": [dayIndex["T."], dayIndex["Sat."]],
      "PLDT ILIGAN": [dayIndex["M."], dayIndex["W."], dayIndex["F."]],
      "SMART GAISANO MALL OZAMIZ": [dayIndex["F."]],
    }
  },

}

Object.freeze(CONFIG);
