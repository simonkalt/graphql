import {
  getJobs,
  getJob,
  getJobsByCompany,
  createJob,
  deleteJob,
  updateJob,
} from "./db/jobs.js";
import { getCompany } from "./db/companies.js";
import { GraphQLError } from "graphql";

export const resolvers = {
  Query: {
    job: async (_root, { id }) => {
      const job = await getJob(id);
      if (!job) {
        throw notFoundError("No Job found with id " + id);
      }
      return job;
    },

    jobs: async () => getJobs(),
    company: async (_root, { id }) => {
      const company = await getCompany(id);
      if (!company) {
        throw notFoundError("No Company found with id " + id);
      }
      return company;
    },
  },

  Mutation: {
    createJob: (_root, { input: { title, description } }, { user }) => {
      // console.log("[createJob] context: ", auth);
      if (!user) {
        throw unauthorizedError("Missing Authentication");
      }
      // console.log("[create job] user: ", user);
      return createJob({ companyId: user.companyId, title, description });
    },
    deleteJob: async (_root, { id }, { user }) => {
      if (!user) {
        throw unauthorizedError("Missing Authentication");
      }
      const job = await deleteJob(id, user.companyId);
      if (!job) {
        throw notFoundError("No Job found with id " + id);
      }
      return job;
    },
    updateJob: async (
      _root,
      { input: { id, title, description } },
      { user }
    ) => {
      if (!user) {
        throw unauthorizedError("Missing Authentication");
      }
      const job = await updateJob({
        id,
        companyId: user.companyId,
        title,
        description,
      });
      if (!job) {
        throw notFoundError("No Job found with id " + id);
      }
    },
  },

  Company: {
    jobs: (company) => getJobsByCompany(company.id),
  },
  Job: {
    date: (job) => toIsoDate(job.createdAt),
    company: (job) => getCompany(job.companyId),
  },
};

function toIsoDate(value) {
  return value.slice(0, "yyyy-mm-dd".length);
}

function notFoundError(message) {
  return new GraphQLError(message, {
    extensions: { code: "NOT_FOUND" },
  });
}

function unauthorizedError(message) {
  return new GraphQLError(message, {
    extensions: { code: "UNAUTHORIZED" },
  });
}
