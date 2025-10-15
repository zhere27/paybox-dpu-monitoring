const CONFIG = {
  APP: {
    NAME: 'Paybox - DPU Monitoring 2.1',
    VERSION: '1.0.0',

  },

  ETAP: {
    SERVICE_BANK: 'eTap',
    ENVIRONMENT: 'testing',
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

  SCHEDULE_BASED: {
    stores: ["PLDT ROBINSONS DUMAGUETE", "SMART SM BACOLOD 1", "SMART SM BACOLOD 2", "SMART SM BACOLOD 3", "PLDT ILIGAN", "SMART GAISANO MALL OZAMIZ"],
    schedules: {
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