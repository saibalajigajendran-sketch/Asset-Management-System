import type { ISharePointField } from './hardwareLaptopSharePointService';

const choicesFor = (field: ISharePointField): string[] => Array.isArray(field.Choices) ? field.Choices : field.Choices?.results || [];

const isEmpty = (value: unknown): boolean => value === undefined || value === null ||
  (typeof value === 'string' && value.trim() === '') ||
  (Array.isArray(value) && value.length === 0);

export interface IAssetFormValidationOptions {
  requireAllFields?: boolean;
  validPersonIdentities?: string[];
}

export const validateAssetForm = (
  fields: ISharePointField[],
  values: Record<string, unknown>,
  options: IAssetFormValidationOptions = {}
): Record<string, string> => {
  const errors: Record<string, string> = {};
  const validPeople = new Set((options.validPersonIdentities || []).map(identity => identity.toLocaleLowerCase()));
  fields.forEach(field => {
    const value = values[field.InternalName];
    const emptyFieldValue = field.TypeAsString === 'MultiChoice' || field.TypeAsString === 'UserMulti'
      ? !String(value ?? '').split(';#').some(item => item.trim())
      : isEmpty(value);
    if ((field.Required || options.requireAllFields) && emptyFieldValue) {
      errors[field.InternalName] = `${field.Title} is required.`;
      return;
    }
    if (isEmpty(value)) return;

    if (options.requireAllFields && (field.TypeAsString === 'User' || field.TypeAsString === 'UserMulti')) {
      const selectedPeople = String(value).split(';#').filter(Boolean);
      if (selectedPeople.some(identity => !validPeople.has(identity.trim().toLocaleLowerCase()))) {
        errors[field.InternalName] = `Select a valid person for ${field.Title}.`;
      }
      return;
    }

    if (field.TypeAsString === 'Choice') {
      const choices = choicesFor(field);
      if (!choices.length) {
        errors[field.InternalName] = `Valid options for ${field.Title} are unavailable. Refresh the register and try again.`;
      } else if (!choices.some(choice => choice.toLocaleLowerCase() === String(value).trim().toLocaleLowerCase())) {
        errors[field.InternalName] = `Select a valid ${field.Title}.`;
      }
    } else if (field.TypeAsString === 'MultiChoice') {
      const choices = choicesFor(field);
      const valuesList = String(value).split(';#').filter(Boolean);
      if (!choices.length) {
        errors[field.InternalName] = `Valid options for ${field.Title} are unavailable. Refresh the register and try again.`;
      } else if (valuesList.some(item => !choices.some(choice => choice.toLocaleLowerCase() === item.trim().toLocaleLowerCase()))) {
        errors[field.InternalName] = `Select valid options for ${field.Title}.`;
      }
    } else if (field.TypeAsString === 'Number' || field.TypeAsString === 'Currency' || field.TypeAsString === 'Integer') {
      if (!Number.isFinite(Number(value))) errors[field.InternalName] = `Enter a valid number for ${field.Title}.`;
    } else if (field.TypeAsString === 'DateTime' && !Number.isFinite(new Date(String(value)).getTime())) {
      errors[field.InternalName] = `Enter a valid date for ${field.Title}.`;
    }
  });
  return errors;
};
